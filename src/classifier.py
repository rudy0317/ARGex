import re
from typing import Any
import cv2
import numpy as np
import pytesseract

DIGITS_7SEG_LOOKUP: dict[tuple[int, ...], str] = {
    (1, 1, 1, 0, 1, 1, 1): "0",
    (1, 0, 1, 1, 1, 0, 1): "2",
    (1, 0, 1, 1, 0, 1, 1): "3",
    (0, 1, 1, 1, 0, 1, 0): "4",
    (1, 1, 0, 1, 0, 1, 1): "5",
    (1, 1, 0, 1, 1, 1, 1): "6",
    (1, 0, 1, 0, 0, 1, 0): "7",
    (1, 1, 1, 0, 0, 1, 0): "7",
    (1, 1, 1, 1, 1, 1, 1): "8",
    (1, 1, 1, 1, 0, 1, 1): "9",
}


class AuditClassifier:
    """Engine klasifikasi visual & OCR berbasis konfigurasi rules.json (tanpa hardcode)."""

    def __init__(self, config: dict[str, Any]) -> None:
        self.config = config
        self.pattern = re.compile(config["location_regex"], re.IGNORECASE | re.MULTILINE)
        self.stencil_map = str.maketrans(config.get("stencil_char_map", {}))
        self.blacklist = {w.upper() for w in config.get("blacklist_prefixes", [])}
        self.detectors = config.get("detectors", {})

    def parse_code(self, text: str) -> str | None:
        """Ekstrak dan normalisasi kode lokasi/DP (contoh: 'LTS\\n~ O10G03' -> 'LTS010603')."""
        if not text:
            return None
        for prefix, raw_digits in self.pattern.findall(text):
            prefix_up = prefix.upper()
            digits_norm = raw_digits.translate(self.stencil_map)
            if digits_norm.isdigit() and prefix_up not in self.blacklist:
                return f"{prefix_up}{digits_norm}"
        return None

    def extract_folder_code(
        self, img: cv2.typing.MatLike, filepath_parts: tuple[str, ...], filename: str
    ) -> str:
        """Cari kode folder lokasi (misal LTS010603) dari hierarki folder, nama file, atau watermark."""
        for part in reversed(filepath_parts[:-1]):
            if code := self.parse_code(part):
                return code

        if code := self.parse_code(filename):
            return code

        h, w, _ = img.shape
        crop_wm = img[int(h * 0.75) : h, int(w * 0.55) : w]
        gray_wm = cv2.cvtColor(crop_wm, cv2.COLOR_BGR2GRAY)
        text_wm = pytesseract.image_to_string(gray_wm, config="--psm 6")
        if code := self.parse_code(text_wm):
            return code

        return "UNKNOWN_LOCATION"

    @staticmethod
    def _decode_7seg_digit(roi_bin: np.ndarray) -> str:
        """Decode 1 karakter angka 7-segment dari citra biner."""
        dh, dw = roi_bin.shape
        if dh == 0 or dw == 0:
            return "?"
        if dw / float(dh) < 0.35:
            return "1"
        segs = [
            ((int(dw * 0.22), 0), (int(dw * 0.78), int(dh * 0.20))),
            ((0, int(dh * 0.12)), (int(dw * 0.35), int(dh * 0.48))),
            ((int(dw * 0.65), int(dh * 0.12)), (dw, int(dh * 0.48))),
            ((int(dw * 0.22), int(dh * 0.40)), (int(dw * 0.78), int(dh * 0.60))),
            ((0, int(dh * 0.52)), (int(dw * 0.35), int(dh * 0.88))),
            ((int(dw * 0.65), int(dh * 0.52)), (dw, int(dh * 0.88))),
            ((int(dw * 0.22), int(dh * 0.80)), (int(dw * 0.78), dh)),
        ]
        state = tuple(
            1 if float(np.mean(roi_bin[y1:y2, x1:x2] > 0)) > 0.35 else 0
            for (x1, y1), (x2, y2) in segs
        )
        return DIGITS_7SEG_LOOKUP.get(state, "0" if state[0] and state[6] else "?")

    def detect_opm_dbm(self, img: cv2.typing.MatLike) -> tuple[str | None, str, bool]:
        """
        Deteksi alat ukur Optical Power Meter (OPM) & baca nilai redaman dBm di layar 7-segment:
        - Redaman < 10 dBm (misal -05.38 dBm) -> 'Input DP'
        - Redaman >= 10 dBm (misal -17.71 dBm) -> 'Output DP'
        Return: (category, detail, is_opm_device)
        """
        rule = self.detectors.get("OPM")
        if not rule:
            return None, "No OPM rule", False

        h, w, _ = img.shape
        roi_cfg = rule["roi"]
        crop = img[
            int(h * roi_cfg["y_start"]) : int(h * roi_cfg["y_end"]),
            int(w * roi_cfg["x_start"]) : int(w * roi_cfg["x_end"]),
        ]
        hsv = cv2.cvtColor(crop, cv2.COLOR_BGR2HSV)
        yb = rule["hsv_yellow_body"]
        yel_mask_crop = (
            (hsv[:, :, 0] >= yb["h_min"])
            & (hsv[:, :, 0] <= yb["h_max"])
            & (hsv[:, :, 1] > yb["s_min"])
            & (hsv[:, :, 2] > yb["v_min"])
        )
        yel_ratio = float(yel_mask_crop.mean())
        if yel_ratio < yb["min_ratio"]:
            return None, f"Not OPM (yel={yel_ratio:.1%})", False

        hsv_full = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        yel_full = (
            (hsv_full[:, :, 0] >= 14)
            & (hsv_full[:, :, 0] <= yb["h_max"])
            & (hsv_full[:, :, 1] > 110)
            & (hsv_full[:, :, 2] > 90)
        ).astype(np.uint8) * 255
        cnts, _ = cv2.findContours(yel_full, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not cnts:
            return None, "OPM contour not found", True

        ox, oy, ow, oh = cv2.boundingRect(max(cnts, key=cv2.contourArea))
        # Pangkas ekor patchcord kuning tipis di bagian atas agar bounding box fokus ke bodi OPM
        sub_yel = yel_full[oy : oy + oh, ox : ox + ow] > 0
        row_y = np.where(sub_yel.mean(axis=1) > 0.18)[0]
        if len(row_y) > 10:
            oy, oh = oy + int(row_y[0]), oh - int(row_y[0])

        opm_top = img[
            oy + int(oh * 0.10) : oy + int(oh * 0.60),
            ox + int(ow * 0.10) : ox + int(ow * 0.90),
        ]
        hsv_top = cv2.cvtColor(opm_top, cv2.COLOR_BGR2HSV)
        backlit = (hsv_top[:, :, 1] < 85) & (hsv_top[:, :, 2] > 150)
        if float(np.mean(backlit)) > 0.15:
            lcd_mask = backlit.astype(np.uint8) * 255
        else:
            ul = rule.get(
                "hsv_unlit_lcd",
                {"h_min": 32, "h_max": 68, "s_min": 18, "s_max": 75, "v_min": 95, "v_max": 155},
            )
            unlit_stn = (
                (hsv_top[:, :, 0] >= ul["h_min"])
                & (hsv_top[:, :, 0] <= ul["h_max"])
                & (hsv_top[:, :, 1] >= ul["s_min"])
                & (hsv_top[:, :, 1] <= ul["s_max"])
                & (hsv_top[:, :, 2] >= ul["v_min"])
                & (hsv_top[:, :, 2] <= ul["v_max"])
            )
            lcd_mask = unlit_stn.astype(np.uint8) * 255
            lcd_mask = cv2.morphologyEx(
                lcd_mask, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (15, 15))
            )
            lcd_mask = cv2.morphologyEx(
                lcd_mask, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (7, 7))
            )

        lcd_cnts, _ = cv2.findContours(lcd_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not lcd_cnts:
            return None, f"OPM detected (yel={yel_ratio:.1%}), LCD screen not found", True

        lcd_cnt = max(lcd_cnts, key=cv2.contourArea)
        pts_src = cv2.boxPoints(cv2.minAreaRect(lcd_cnt)).astype("float32")
        s, diff = pts_src.sum(axis=1), np.diff(pts_src, axis=1)
        ordered = np.array(
            [
                pts_src[np.argmin(s)],
                pts_src[np.argmin(diff)],
                pts_src[np.argmax(s)],
                pts_src[np.argmax(diff)],
            ],
            dtype="float32",
        )
        W, H = 320, 220
        M = cv2.getPerspectiveTransform(
            ordered, np.array([[0, 0], [W - 1, 0], [W - 1, H - 1], [0, H - 1]], dtype="float32")
        )
        warped = cv2.warpPerspective(opm_top, M, (W, H))

        dbm_row = warped[int(H * 0.50) : int(H * 0.96), int(W * 0.02) : int(W * 0.78)]
        gray = cv2.cvtColor(dbm_row, cv2.COLOR_BGR2GRAY)
        _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
        closed = cv2.morphologyEx(
            th, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (4, 12))
        )
        d_cnts, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        boxes = sorted(
            [
                cv2.boundingRect(c)
                for c in d_cnts
                if cv2.boundingRect(c)[3] > 40 and cv2.boundingRect(c)[0] > 10
            ],
            key=lambda b: b[0],
        )

        if len(boxes) >= 2:
            digits = [
                self._decode_7seg_digit(th[y : y + hb, x : x + wb])
                for (x, y, wb, hb) in boxes[:4]
            ]
            tens_digit = digits[0]
            raw_str = "".join(digits[:2]) + ("." + "".join(digits[2:4]) if len(digits) >= 4 else "")
            try:
                abs_dbm = float(raw_str.replace("?", "0"))
            except ValueError:
                abs_dbm = 5.0 if tens_digit == "0" else 15.0

            sop_limit = float(rule.get("sop_dbm_threshold", 10.0))
            if abs_dbm < sop_limit:
                return (
                    "Input DP",
                    f"OPM -{raw_str} dBm (< {sop_limit} dBm SOP Input, yel={yel_ratio:.1%})",
                    True,
                )
            return (
                "Output DP",
                f"OPM -{raw_str} dBm (>= {sop_limit} dBm SOP Output, yel={yel_ratio:.1%})",
                True,
            )

        return None, f"OPM detected (yel={yel_ratio:.1%}), gagal baca digit LCD", True

    def detect_basetray_dp(self, img: cv2.typing.MatLike, rule: dict[str, Any]) -> tuple[bool, str]:
        """Deteksi 'Basetray DP' (Nomor 6): Tray putih menghadap depan + deretan port SC biru."""
        h, w, _ = img.shape
        roi = rule["roi"]
        c = img[
            int(h * roi["y_start"]) : int(h * roi["y_end"]),
            int(w * roi["x_start"]) : int(w * roi["x_end"]),
        ]
        hsv = cv2.cvtColor(c, cv2.COLOR_BGR2HSV)
        mid_x = int(c.shape[1] * 0.45)

        blue_sc = float(
            np.mean((hsv[:, :, 0] >= 98) & (hsv[:, :, 0] <= 128) & (hsv[:, :, 1] > 110) & (hsv[:, :, 2] > 70))
        )
        white_tray = float(np.mean((hsv[:, :, 1] < 45) & (hsv[:, :, 2] > 165)))
        left_dark = float(np.mean(hsv[:, :mid_x, 2] < 90))
        left_orange = float(
            np.mean((hsv[:, :mid_x, 0] >= 5) & (hsv[:, :mid_x, 0] <= 18) & (hsv[:, :mid_x, 1] > 120) & (hsv[:, :mid_x, 2] > 100))
        )

        if (
            blue_sc >= rule["min_blue_sc_ratio"]
            and white_tray >= rule["min_white_tray_ratio"]
            and left_dark <= rule["max_left_dark_ratio"]
            and left_orange <= rule.get("max_left_orange_ratio", 0.005)
        ):
            return (
                True,
                f"Match Basetray DP (blueSC={blue_sc:.1%}, whiteTray={white_tray:.1%}, leftOrg={left_orange:.1%})",
            )
        return False, f"Not Basetray (blueSC={blue_sc:.1%}, white={white_tray:.1%}, leftOrg={left_orange:.1%})"

    def detect_loop_dp(self, img: cv2.typing.MatLike, rule: dict[str, Any]) -> tuple[bool, str]:
        """Deteksi 'Loop DP' (Nomor 3): Pintu kiri hitam dengan loop tube biru/oranye + kanan tray putih."""
        h, w, _ = img.shape
        roi = rule["roi"]
        c = img[
            int(h * roi["y_start"]) : int(h * roi["y_end"]),
            int(w * roi["x_start"]) : int(w * roi["x_end"]),
        ]
        hsv = cv2.cvtColor(c, cv2.COLOR_BGR2HSV)
        mid_x = int(c.shape[1] * 0.45)

        left_dark = float(np.mean(hsv[:, :mid_x, 2] < 90))
        right_white = float(np.mean((hsv[:, mid_x:, 1] < 45) & (hsv[:, mid_x:, 2] > 155)))
        blue_tube = float(
            np.mean((hsv[:, :, 0] >= 98) & (hsv[:, :, 0] <= 128) & (hsv[:, :, 1] > 110) & (hsv[:, :, 2] > 70))
        )
        orange_tube = float(
            np.mean((hsv[:, :, 0] >= 5) & (hsv[:, :, 0] <= 18) & (hsv[:, :, 1] > 120) & (hsv[:, :, 2] > 100))
        )

        if (
            left_dark >= rule["min_left_dark_ratio"]
            and right_white >= rule["min_right_white_ratio"]
            and blue_tube >= rule["min_blue_tube_ratio"]
            and orange_tube >= rule["min_orange_tube_ratio"]
        ):
            return (
                True,
                f"Match Loop DP (leftDark={left_dark:.1%}, rightWhite={right_white:.1%}, blue={blue_tube:.1%}, org={orange_tube:.1%})",
            )
        return False, f"Not Loop DP (leftDark={left_dark:.1%}, rightWhite={right_white:.1%}, blue={blue_tube:.1%}, org={orange_tube:.1%})"

    def detect_dp(self, img: cv2.typing.MatLike, rule: dict[str, Any]) -> tuple[bool, str]:
        """Evaluasi apakah foto masuk kategori 'DP' berdasarkan parameter di rules.json."""
        h, w, _ = img.shape
        roi = rule["roi"]
        crop_center = img[
            int(h * roi["y_start"]) : int(h * roi["y_end"]),
            int(w * roi["x_start"]) : int(w * roi["x_end"]),
        ]

        hsv = cv2.cvtColor(crop_center, cv2.COLOR_BGR2HSV)

        if excl := rule["hsv_gates"].get("exclude_opm_yellow"):
            yel_mask = (
                (hsv[:, :, 0] >= excl["h_min"])
                & (hsv[:, :, 0] <= excl["h_max"])
                & (hsv[:, :, 1] > excl["s_min"])
                & (hsv[:, :, 2] > excl["v_min"])
            )
            if float(yel_mask.mean()) > excl["max_ratio"]:
                return False, "Excluded (OPM yellow detected)"

        g_black = rule["hsv_gates"]["black_box"]
        dark_mask = (
            (hsv[:, :, 1] < g_black["s_max"])
            & (hsv[:, :, 2] >= g_black["v_min"])
            & (hsv[:, :, 2] <= g_black["v_max"])
        )
        dark_ratio = float(dark_mask.mean())

        g_white = rule["hsv_gates"]["white_stencil"]
        white_mask = (hsv[:, :, 1] < g_white["s_max"]) & (hsv[:, :, 2] > g_white["v_min"])
        white_ratio = float(white_mask.mean())

        pass_black = dark_ratio >= g_black["min_ratio"]
        pass_white = g_white["min_ratio"] <= white_ratio <= g_white["max_ratio"]

        if not (pass_black and pass_white):
            return (
                False,
                f"Skip HSV Gate (dark={dark_ratio:.1%}, white={white_ratio:.1%})",
            )

        if not rule.get("require_ocr_match", True):
            return True, f"Match HSV Gate (dark={dark_ratio:.1%}, white={white_ratio:.1%})"

        gray = cv2.cvtColor(crop_center, cv2.COLOR_BGR2GRAY)
        candidates = []
        for thresh_val in rule.get("ocr_thresholds", [170, 150]):
            _, th = cv2.threshold(gray, thresh_val, 255, cv2.THRESH_BINARY)
            candidates.append(th)
        _, th_otsu = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        candidates.extend([th_otsu, gray])

        for proc in candidates:
            extracted_text = pytesseract.image_to_string(proc, config="--psm 6")
            if dp_code := self.parse_code(extracted_text):
                return (
                    True,
                    f"Match DP '{dp_code}' (dark={dark_ratio:.1%}, white={white_ratio:.1%})",
                )

        return (
            False,
            f"Box hitam cocok (dark={dark_ratio:.1%}), namun teks sablon gagal dibaca OCR",
        )

    def detect_structure(self, img: cv2.typing.MatLike, rule: dict[str, Any]) -> tuple[bool, str]:
        """Deteksi 'Structure' (Nomor 8): Foto tiang vertikal utuh dari jalan ke atas."""
        h, w, _ = img.shape
        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # Area atas (langit cerah atau pepohonan hijau di y: 0.05..0.45)
        top_roi = hsv[int(h * 0.05) : int(h * 0.45), :]
        top_sky = (top_roi[:, :, 1] < 45) & (top_roi[:, :, 2] > 175)
        top_green = (
            (top_roi[:, :, 0] >= 35)
            & (top_roi[:, :, 0] <= 85)
            & (top_roi[:, :, 1] > 40)
            & (top_roi[:, :, 2] > 35)
        )
        top_outdoor = float(np.mean(top_sky | top_green))

        # Area tengah-bawah (harus berupa jalan/bangunan/pohon, BUKAN langit mendongak atau cor beton terang)
        mid_low = hsv[int(h * 0.45) : int(h * 0.75), int(w * 0.20) : int(w * 0.80)]
        mid_low_bright = float(np.mean((mid_low[:, :, 1] < 40) & (mid_low[:, :, 2] > 165)))

        # Dominasi garis vertikal tiang di jalur tengah
        sobelx = np.abs(
            cv2.Sobel(
                gray[int(h * 0.15) : int(h * 0.80), int(w * 0.30) : int(w * 0.70)],
                cv2.CV_64F,
                1,
                0,
                ksize=3,
            )
        )
        vert_edge = float(np.mean(sobelx > 50))

        if (
            top_outdoor >= rule["min_top_outdoor_ratio"]
            and mid_low_bright <= rule["max_mid_low_bright_ratio"]
            and vert_edge >= rule["min_vertical_edge_ratio"]
        ):
            return (
                True,
                f"Match Structure (topOutdoor={top_outdoor:.1%}, midLowBright={mid_low_bright:.1%}, vertEdge={vert_edge:.1%})",
            )
        return (
            False,
            f"Not Structure (topOutdoor={top_outdoor:.1%}, midLowBright={mid_low_bright:.1%})",
        )

    def classify(
        self, img: cv2.typing.MatLike, active_targets: list[str]
    ) -> tuple[str | None, str]:
        """Cek gambar terhadap daftar target variabel yang sedang aktif."""
        last_reason = "No active detector matched"

        # 1. OPM (Input DP / Output DP)
        if "Input DP" in active_targets or "Output DP" in active_targets:
            opm_cat, opm_detail, is_opm_device = self.detect_opm_dbm(img)
            if opm_cat and opm_cat in active_targets:
                return opm_cat, opm_detail
            if is_opm_device:
                return None, opm_detail
            last_reason = opm_detail

        # 2. Basetray DP (Nomor 6)
        if "Basetray DP" in active_targets and "Basetray DP" in self.detectors:
            matched, detail = self.detect_basetray_dp(img, self.detectors["Basetray DP"])
            if matched:
                return "Basetray DP", detail
            last_reason = detail

        # 3. Loop DP (Nomor 3)
        if "Loop DP" in active_targets and "Loop DP" in self.detectors:
            matched, detail = self.detect_loop_dp(img, self.detectors["Loop DP"])
            if matched:
                return "Loop DP", detail
            last_reason = detail

        # 4. DP (Nomor 1)
        if "DP" in active_targets and "DP" in self.detectors:
            matched, detail = self.detect_dp(img, self.detectors["DP"])
            if matched:
                return "DP", detail
            last_reason = detail

        # 5. Structure (Nomor 8)
        if "Structure" in active_targets and "Structure" in self.detectors:
            matched, detail = self.detect_structure(img, self.detectors["Structure"])
            if matched:
                return "Structure", detail
            last_reason = detail

        return None, last_reason
