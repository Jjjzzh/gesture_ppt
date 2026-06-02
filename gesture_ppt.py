import cv2
import numpy as np
import time
import os
import sys
import shutil
from collections import deque

import mediapipe as mp
from mediapipe.tasks.python import vision
from mediapipe.tasks.python.core import base_options as bo


def _get_ppt_app():
    """Connect to a running PowerPoint instance or launch one."""
    try:
        from win32com.client import Dispatch, GetActiveObject
        ppt = GetActiveObject("PowerPoint.Application")
        ppt.Visible = True
        return ppt
    except Exception:
        ppt = Dispatch("PowerPoint.Application")
        ppt.Visible = True
        return ppt

# ── Configuration ──────────────────────────────────────────────
SWIPE_THRESHOLD = 0.06
SWIPE_FRAMES = 8
FIST_HOLD_FRAMES = 8
THUMB_HOLD_FRAMES = 5
COOLDOWN_SEC = 1.2
TRAIL_LENGTH = 20
CAMERA_W, CAMERA_H = 1280, 720

# Window placement (corner PIP)
WINDOW_W, WINDOW_H = 120,80
WINDOW_POSITION = "top-right"   # top-left | top-right | bottom-left | bottom-right
WINDOW_TOPMOST = True

# Model must be at an ASCII path (MediaPipe C library limitation)
MODEL_PATH = "C:/temp/hand_landmarker.task"

# colors (BGR)
NEON_GREEN = (0, 255, 100)
NEON_CYAN = (255, 200, 0)
NEON_PINK = (180, 0, 255)
NEON_GOLD = (0, 215, 255)
WHITE = (255, 255, 255)

# MediaPipe hand skeleton connections (start, end landmark indices)
HAND_CONNECTIONS = [
    (0, 1), (1, 2), (2, 3), (3, 4),
    (0, 5), (5, 6), (6, 7), (7, 8),
    (5, 9), (9, 10), (10, 11), (11, 12),
    (9, 13), (13, 14), (14, 15), (15, 16),
    (13, 17), (17, 18), (18, 19), (19, 20),
    (0, 17),
]


class GestureDetector:
    def __init__(self, model_path):
        options = vision.HandLandmarkerOptions(
            base_options=bo.BaseOptions(model_asset_path=model_path),
            running_mode=vision.RunningMode.VIDEO,
            num_hands=1,
            min_hand_detection_confidence=0.7,
            min_hand_presence_confidence=0.7,
            min_tracking_confidence=0.7,
        )
        self.detector = vision.HandLandmarker.create_from_options(options)
        self.trail = deque(maxlen=TRAIL_LENGTH)
        self.counter = 0
        self._prev_gesture = None
        self.current_gesture = "none"
        self.frame_ts = 0

    def process(self, frame):
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        mp_image = mp.Image(image_format=mp.ImageFormat.SRGB, data=rgb)
        self.frame_ts += 1
        result = self.detector.detect_for_video(mp_image, self.frame_ts)

        self.current_gesture = "none"
        hand_landmarks = None

        if result.hand_landmarks and len(result.hand_landmarks) > 0:
            lm_list = result.hand_landmarks[0]
            hand_landmarks = lm_list
            norm = [(lm.x, lm.y) for lm in lm_list]

            gesture = self._classify(norm)
            self.current_gesture = gesture

            palm = ((norm[0][0] + norm[9][0]) / 2, (norm[0][1] + norm[9][1]) / 2)
            self.trail.append(palm)

            # hold counter: accumulate for same gesture, reset on change
            if gesture == self._prev_gesture:
                self.counter += 1
            else:
                self.counter = 1
            self._prev_gesture = gesture
        else:
            self.trail.clear()
            self.counter = 0
            self._prev_gesture = None

        return hand_landmarks

    def _classify(self, norm):
        """Classify: open_palm / one / two / none."""
        # check each finger extended (tip above PIP)
        idx_up = norm[8][1] < norm[6][1]
        mid_up = norm[12][1] < norm[10][1]
        ring_up = norm[16][1] < norm[14][1]
        pinky_up = norm[20][1] < norm[18][1]
        thumb_out = abs(norm[4][0] - norm[3][0]) > 0.04
        fingers_up = idx_up + mid_up + ring_up + pinky_up

        # open palm: all 5 fingers spread
        if fingers_up >= 3 and thumb_out:
            return "open_palm"

        # 比1: only index extended, others curled
        if idx_up and not mid_up and not ring_up and not pinky_up:
            return "one"

        # 比2: index + middle extended, ring + pinky curled
        if idx_up and mid_up and not ring_up and not pinky_up:
            return "two"

        return "none"

    def should_start_slideshow(self):
        return self.current_gesture == "open_palm" and self.counter >= FIST_HOLD_FRAMES

    def should_next_slide(self):
        return self.current_gesture == "one" and self.counter >= THUMB_HOLD_FRAMES

    def should_prev_slide(self):
        return self.current_gesture == "two" and self.counter >= THUMB_HOLD_FRAMES

    def reset_counter(self):
        self.counter = 0

    def close(self):
        self.detector.close()


class VisualOverlay:
    def __init__(self):
        self.flash_alpha = 0.0
        self.flash_color = (0, 0, 0)
        self.flash_text = ""
        self.flash_text_alpha = 0.0
        self.slide_index = 1
        self.fps_history = deque(maxlen=30)
        self.prev_tick = time.time()

    def trigger_flash(self, color, text):
        self.flash_alpha = 1.0
        self.flash_color = color
        self.flash_text = text
        self.flash_text_alpha = 1.0

    def update_slide(self, direction):
        if direction == "next":
            self.slide_index += 1
        elif direction == "prev":
            self.slide_index = max(1, self.slide_index - 1)
        elif direction == "start":
            self.slide_index = 1

    def draw(self, frame, hand_landmarks, gesture, trail, fps):
        h, w = frame.shape[:2]
        overlay = frame.copy()
        overlay = cv2.addWeighted(overlay, 0.75, np.zeros_like(overlay), 0.25, 0)

        # ── swipe trail ──
        if len(trail) > 1:
            for i in range(len(trail) - 1):
                alpha = (i + 1) / len(trail)
                pt1 = (int(trail[i][0] * w), int(trail[i][1] * h))
                pt2 = (int(trail[i + 1][0] * w), int(trail[i + 1][1] * h))
                c = (int(NEON_CYAN[0] * alpha), int(NEON_CYAN[1] * alpha), int(NEON_CYAN[2] * alpha))
                cv2.line(overlay, pt1, pt2, c, 2, cv2.LINE_AA)
                if i % 3 == 0:
                    cv2.circle(overlay, pt1, int(6 * alpha), c, -1, cv2.LINE_AA)

        # ── hand landmarks ──
        if hand_landmarks is not None and len(hand_landmarks) > 0:
            if gesture == "open_palm":
                lc, cc = (255, 105, 180), (220, 60, 140)
            elif gesture == "one":
                lc, cc = (0, 255, 180), (0, 210, 140)
            elif gesture == "two":
                lc, cc = (255, 200, 80), (220, 160, 40)
            else:
                lc, cc = NEON_PINK, (140, 0, 220)

            for lm in hand_landmarks:
                cx, cy = int(lm.x * w), int(lm.y * h)
                cv2.circle(overlay, (cx, cy), 8, lc, -1, cv2.LINE_AA)
                cv2.circle(overlay, (cx, cy), 4, WHITE, -1, cv2.LINE_AA)

            for s, e in HAND_CONNECTIONS:
                if s < len(hand_landmarks) and e < len(hand_landmarks):
                    p1, p2 = hand_landmarks[s], hand_landmarks[e]
                    cv2.line(overlay, (int(p1.x * w), int(p1.y * h)),
                             (int(p2.x * w), int(p2.y * h)), cc, 2, cv2.LINE_AA)

            if trail:
                cx, cy = int(trail[-1][0] * w), int(trail[-1][1] * h)
                pulse = int(15 + 5 * np.sin(time.time() * 6))
                cv2.circle(overlay, (cx, cy), pulse, NEON_CYAN, 2, cv2.LINE_AA)

        # ── HUD panel (top-left) ──
        px, py, pw, ph = 30, 30, 260, 200
        self._rounded_rect(overlay, px, py, pw, ph, (15, 15, 15), 0.75, 16, -1)
        self._rounded_rect(overlay, px, py, pw, ph, NEON_CYAN, 0.3, 16, 2)

        names = {"open_palm": "Open Palm", "one": "One",
                 "two": "Two", "none": "Scanning..."}
        colors = {"open_palm": (255, 105, 180), "one": (0, 255, 180),
                  "two": (255, 200, 80), "none": NEON_PINK}
        gname = names.get(gesture, "Unknown")
        gcolor = colors.get(gesture, WHITE)

        cv2.putText(overlay, "GESTURE", (px + 20, py + 35),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.55, (150, 150, 150), 1, cv2.LINE_AA)
        cv2.putText(overlay, gname, (px + 20, py + 75),
                    cv2.FONT_HERSHEY_SIMPLEX, 1.2, gcolor, 2, cv2.LINE_AA)
        cv2.putText(overlay, "SLIDE", (px + 20, py + 115),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.55, (150, 150, 150), 1, cv2.LINE_AA)
        cv2.putText(overlay, f"#{self.slide_index}", (px + 20, py + 153),
                    cv2.FONT_HERSHEY_SIMPLEX, 1.1, WHITE, 2, cv2.LINE_AA)

        hy = py + ph + 20
        self._rounded_rect(overlay, px, hy, pw, 110, (15, 15, 15), 0.75, 16, -1)
        self._rounded_rect(overlay, px, hy, pw, 110, NEON_GREEN, 0.2, 16, 2)
        cv2.putText(overlay, "CONTROLS", (px + 20, hy + 28),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (150, 150, 150), 1, cv2.LINE_AA)
        cv2.putText(overlay, "Open Palm   >  Start", (px + 20, hy + 55),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (200, 200, 200), 1, cv2.LINE_AA)
        cv2.putText(overlay, "One (1)      >  Next", (px + 20, hy + 78),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (200, 200, 200), 1, cv2.LINE_AA)
        cv2.putText(overlay, "Two (2)      >  Prev", (px + 20, hy + 101),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (200, 200, 200), 1, cv2.LINE_AA)

        # ── FPS (top-right) ──
        fps_text = f"FPS: {fps:.0f}"
        (tw, th), _ = cv2.getTextSize(fps_text, cv2.FONT_HERSHEY_SIMPLEX, 0.55, 2)
        self._rounded_rect(overlay, w - tw - 50, 20, tw + 40, 40, (15, 15, 15), 0.75, 10, -1)
        cv2.putText(overlay, fps_text, (w - tw - 30, 47),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.55, NEON_CYAN, 2, cv2.LINE_AA)

        # ── action flash ──
        if self.flash_alpha > 0.001:
            fo = np.zeros_like(frame)
            fo[:, :] = self.flash_color
            overlay = cv2.addWeighted(overlay, 1.0, fo, self.flash_alpha * 0.25, 0)
            if self.flash_text_alpha > 0.001:
                (tw2, th2), _ = cv2.getTextSize(self.flash_text, cv2.FONT_HERSHEY_DUPLEX, 2.5, 4)
                tx, ty = (w - tw2) // 2, h // 2
                cv2.putText(overlay, self.flash_text, (tx, ty),
                            cv2.FONT_HERSHEY_DUPLEX, 2.5, WHITE, 4, cv2.LINE_AA)
            self.flash_alpha *= 0.88
            self.flash_text_alpha *= 0.88

        return overlay

    def _rounded_rect(self, img, x, y, w, h, color, alpha, radius, thickness):
        ov = img.copy()
        cv2.rectangle(ov, (x + radius, y), (x + w - radius, y + h), color, thickness, cv2.LINE_AA)
        cv2.rectangle(ov, (x, y + radius), (x + w, y + h - radius), color, thickness, cv2.LINE_AA)
        if thickness == -1:
            cv2.circle(ov, (x + radius, y + radius), radius, color, -1, cv2.LINE_AA)
            cv2.circle(ov, (x + w - radius, y + radius), radius, color, -1, cv2.LINE_AA)
            cv2.circle(ov, (x + radius, y + h - radius), radius, color, -1, cv2.LINE_AA)
            cv2.circle(ov, (x + w - radius, y + h - radius), radius, color, -1, cv2.LINE_AA)
        else:
            cv2.ellipse(ov, (x + radius, y + radius), (radius, radius), 180, 0, 90, color, thickness, cv2.LINE_AA)
            cv2.ellipse(ov, (x + w - radius, y + radius), (radius, radius), 270, 0, 90, color, thickness, cv2.LINE_AA)
            cv2.ellipse(ov, (x + radius, y + h - radius), (radius, radius), 90, 0, 90, color, thickness, cv2.LINE_AA)
            cv2.ellipse(ov, (x + w - radius, y + h - radius), (radius, radius), 0, 0, 90, color, thickness, cv2.LINE_AA)
        cv2.addWeighted(ov, alpha, img, 1 - alpha, 0, img)


class PPTController:
    """Direct PowerPoint control via COM (no keyboard simulation, no focus issue)."""

    def __init__(self):
        self.last_action_time = 0
        self.slide_index = 1
        self.ppt = None
        self._connected = False
        self._try_connect()

    def _try_connect(self):
        try:
            self.ppt = _get_ppt_app()
            self._connected = True
            print("[PPT] Connected to PowerPoint")
        except Exception as e:
            self._connected = False
            print(f"[PPT] PowerPoint not available: {e}")
            print("[PPT] Please open PowerPoint manually first.")

    def next_slide(self):
        if not self._can_act() or not self._connected:
            return False
        try:
            view = self.ppt.ActivePresentation.SlideShowWindow.View
            view.Next()
            self.slide_index += 1
            self._mark_action()
            return True
        except Exception:
            self._connected = False
            return False

    def prev_slide(self):
        if not self._can_act() or not self._connected:
            return False
        try:
            view = self.ppt.ActivePresentation.SlideShowWindow.View
            view.Previous()
            self.slide_index = max(1, self.slide_index - 1)
            self._mark_action()
            return True
        except Exception:
            self._connected = False
            return False

    def start_slideshow(self):
        if not self._can_act() or not self._connected:
            return False
        try:
            pres = self.ppt.ActivePresentation
            # check if slideshow is already running
            try:
                _ = pres.SlideShowWindow
                return False  # already running
            except Exception:
                pass
            pres.SlideShowSettings.Run()
            self.slide_index = 1
            self._mark_action()
            return True
        except Exception:
            self._connected = False
            return False

    def _can_act(self):
        return time.time() - self.last_action_time > COOLDOWN_SEC

    def _mark_action(self):
        self.last_action_time = time.time()


def blur_background(frame, hand_landmarks, blur_k=45):
    """Blur the entire frame except the hand region (depth-of-field effect)."""
    h, w = frame.shape[:2]
    # ensure odd kernel size
    if blur_k % 2 == 0:
        blur_k += 1
    blurred = cv2.GaussianBlur(frame, (blur_k, blur_k), 0)

    if hand_landmarks is None or len(hand_landmarks) == 0:
        return blurred

    # convex hull of hand landmarks
    pts = np.array([(int(lm.x * w), int(lm.y * h)) for lm in hand_landmarks], dtype=np.int32)
    hull = cv2.convexHull(pts)
    mask = np.zeros((h, w), dtype=np.uint8)
    cv2.fillConvexPoly(mask, hull, 255)

    # expand mask to include full hand + wrist
    kernel = np.ones((35, 35), np.uint8)
    mask = cv2.dilate(mask, kernel, iterations=1)
    mask = cv2.GaussianBlur(mask, (35, 35), 0)  # feather edges

    mask_3ch = mask[:, :, np.newaxis] / 255.0
    return (frame * mask_3ch + blurred * (1 - mask_3ch)).astype(np.uint8)


def draw_hold_progress(img, progress, w, h, gesture):
    bar_w, bar_h = 300, 8
    bx, by = (w - bar_w) // 2, h - 70
    cv2.rectangle(img, (bx, by), (bx + bar_w, by + bar_h), (60, 60, 60), -1, cv2.LINE_AA)
    fill_w = int(bar_w * progress)
    if fill_w > 0:
        c = NEON_GOLD if progress < 1.0 else NEON_GREEN
        cv2.rectangle(img, (bx, by), (bx + fill_w, by + bar_h), c, -1, cv2.LINE_AA)
    labels = {
        "open_palm": "Open Palm -> Start",
        "one": "One -> Next",
        "two": "Two -> Prev",
    }
    label = labels.get(gesture, "")
    if progress >= 1.0:
        label = {"open_palm": "Starting!", "one": "Next!", "two": "Prev!"}.get(gesture, "!")
    (lw, lh), _ = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 0.55, 2)
    cv2.putText(img, label, ((w - lw) // 2, by - 12),
                cv2.FONT_HERSHEY_SIMPLEX, 0.55, WHITE, 2, cv2.LINE_AA)


def ensure_model():
    """Ensure the hand landmarker model file exists at an ASCII path."""
    if os.path.exists(MODEL_PATH):
        return True

    home_model = os.path.expanduser("~/hand_landmarker.task")
    if os.path.exists(home_model):
        os.makedirs(os.path.dirname(MODEL_PATH), exist_ok=True)
        try:
            shutil.copy(home_model, MODEL_PATH)
            print(f"Model copied to {MODEL_PATH}")
            return True
        except Exception as e:
            print(f"Failed to copy model: {e}")

    print(f"ERROR: Model file not found at {MODEL_PATH}")
    print("Download it from:")
    print("  https://storage.googleapis.com/mediapipe-models/hand_landmarker/")
    print("  hand_landmarker/float16/latest/hand_landmarker.task")
    print(f"and save to {MODEL_PATH}")
    return False


def main():
    if not ensure_model():
        sys.exit(1)

    cap = cv2.VideoCapture(0)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, CAMERA_W)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_H)

    detector = GestureDetector(MODEL_PATH)
    overlay_renderer = VisualOverlay()
    controller = PPTController()

    print("=" * 55)
    print("  Gesture PPT Controller  (COM direct control)")
    print("  Open Palm (5) -> Start Slideshow")
    print("  One (1)       -> Next Slide")
    print("  Two (2)       -> Previous Slide")
    print("  Press Q to quit")
    print("=" * 55)

    cv2.namedWindow("Gesture PPT Controller", cv2.WINDOW_NORMAL)
    cv2.resizeWindow("Gesture PPT Controller", WINDOW_W, WINDOW_H)

    # pin to screen corner + always on top
    if WINDOW_TOPMOST:
        cv2.setWindowProperty("Gesture PPT Controller", cv2.WND_PROP_TOPMOST, 1)
    try:
        import ctypes
        screen_w = ctypes.windll.user32.GetSystemMetrics(0)
        screen_h = ctypes.windll.user32.GetSystemMetrics(1)
    except Exception:
        screen_w, screen_h = 1920, 1080
    positions = {
        "top-left":     (20, 40),
        "top-right":    (screen_w - WINDOW_W - 20, 40),
        "bottom-left":  (20, screen_h - WINDOW_H - 60),
        "bottom-right": (screen_w - WINDOW_W - 20, screen_h - WINDOW_H - 60),
    }
    wx, wy = positions.get(WINDOW_POSITION, (screen_w - WINDOW_W - 20, 40))
    cv2.moveWindow("Gesture PPT Controller", wx, wy)

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        frame = cv2.flip(frame, 1)
        h, w = frame.shape[:2]

        hand_landmarks = detector.process(frame)
        gesture = detector.current_gesture

        now = time.time()
        overlay_renderer.fps_history.append(1.0 / max(now - overlay_renderer.prev_tick, 0.001))
        overlay_renderer.prev_tick = now
        fps = np.mean(overlay_renderer.fps_history)

        action_triggered = None

        if detector.should_start_slideshow():
            if controller.start_slideshow():
                action_triggered = ("start", NEON_GOLD)
                overlay_renderer.update_slide("start")
                detector.reset_counter()
                print("[ACTION] Slideshow Started!")

        if detector.should_next_slide():
            if controller.next_slide():
                action_triggered = ("next", NEON_GREEN)
                overlay_renderer.update_slide("next")
                detector.reset_counter()
                print(f"[ACTION] Next Slide -> #{controller.slide_index}")

        if detector.should_prev_slide():
            if controller.prev_slide():
                action_triggered = ("prev", NEON_CYAN)
                overlay_renderer.update_slide("prev")
                detector.reset_counter()
                print(f"[ACTION] Previous Slide -> #{controller.slide_index}")

        if action_triggered:
            overlay_renderer.trigger_flash(action_triggered[1], action_triggered[0].upper())

        # blur background, keep hand sharp
        frame = blur_background(frame, hand_landmarks)
        output = overlay_renderer.draw(frame, hand_landmarks, gesture, detector.trail, fps)

        if gesture in ("open_palm", "one", "two"):
            total = FIST_HOLD_FRAMES if gesture == "open_palm" else THUMB_HOLD_FRAMES
            progress = min(detector.counter / max(total, 1), 1.0)
            draw_hold_progress(output, progress, w, h, gesture)

        cv2.imshow("Gesture PPT Controller", output)

        key = cv2.waitKey(1) & 0xFF
        if key == ord("q") or key == 27:
            break

    detector.close()
    cap.release()
    cv2.destroyAllWindows()
    print("Exit.")


if __name__ == "__main__":
    main()
