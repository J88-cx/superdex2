"""Fixed multi-finger gesture recognition for Superdex.

The recognizer keeps the GestureSign-inspired core small and explicit:
track sampling, coordinate normalization, shape matching, threshold checks, and
mapping a fixed set of recognized gestures to host actions.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

Point = Tuple[float, float]
Stroke = List[Point]
Strokes = List[Stroke]

DIRECTION_TAP = "tap"
DIRECTION_UP = "up"
DIRECTION_DOWN = "down"
DIRECTION_LEFT = "left"
DIRECTION_RIGHT = "right"


@dataclass(frozen=True)
class GestureAction:
    """Concrete action to execute after a gesture is recognized."""

    outer_name: str
    trigger: str
    action_type: str
    payload: Dict[str, object] = field(default_factory=dict)


@dataclass(frozen=True)
class FixedGesture:
    """A fixed gesture template with a required finger count and action."""

    name: str
    finger_count: int
    direction: str
    action: GestureAction
    min_score: float = 60.0


@dataclass(frozen=True)
class GestureMatch:
    """Recognition result returned by FixedGestureRecognizer."""

    gesture: FixedGesture
    score: float

    @property
    def action(self) -> GestureAction:
        return self.gesture.action


class PointPatternAnalyzer:
    """Resample, normalize, and compare stroke shapes."""

    def __init__(self, sample_size: int = 32):
        self.sample_size = max(4, int(sample_size))

    def centroid_stroke(self, strokes: Strokes) -> Stroke:
        """Convert same-time multi-finger strokes to one representative stroke."""
        cleaned = [stroke for stroke in strokes if stroke]
        if not cleaned:
            return []

        sampled = [self.resample(stroke, self.sample_size) for stroke in cleaned]
        return [
            (
                sum(stroke[index][0] for stroke in sampled) / len(sampled),
                sum(stroke[index][1] for stroke in sampled) / len(sampled),
            )
            for index in range(self.sample_size)
        ]

    def normalize(self, stroke: Sequence[Point]) -> Stroke:
        if not stroke:
            return []

        sampled = self.resample(stroke, self.sample_size)
        min_x = min(point[0] for point in sampled)
        min_y = min(point[1] for point in sampled)
        max_x = max(point[0] for point in sampled)
        max_y = max(point[1] for point in sampled)
        width = max(max_x - min_x, 1e-6)
        height = max(max_y - min_y, 1e-6)
        scale = max(width, height)
        return [((x - min_x) / scale, (y - min_y) / scale) for x, y in sampled]

    def resample(self, points: Sequence[Point], target_count: int) -> Stroke:
        if not points:
            return []
        if len(points) == 1:
            return [tuple(points[0]) for _ in range(target_count)]

        target_count = max(2, int(target_count))
        points = [tuple(map(float, point)) for point in points]
        distances = [0.0]
        for previous, current in zip(points, points[1:]):
            distances.append(distances[-1] + self.distance(previous, current))

        total_length = distances[-1]
        if total_length <= 1e-9:
            return [points[0] for _ in range(target_count)]

        step = total_length / (target_count - 1)
        resampled = [points[0]]
        segment_index = 1
        for sample_index in range(1, target_count - 1):
            target_distance = step * sample_index
            while segment_index < len(points) and distances[segment_index] < target_distance:
                segment_index += 1

            if segment_index >= len(points):
                resampled.append(points[-1])
                continue

            previous_point = points[segment_index - 1]
            next_point = points[segment_index]
            previous_distance = distances[segment_index - 1]
            segment_length = max(distances[segment_index] - previous_distance, 1e-9)
            ratio = (target_distance - previous_distance) / segment_length
            resampled.append(
                (
                    previous_point[0] + (next_point[0] - previous_point[0]) * ratio,
                    previous_point[1] + (next_point[1] - previous_point[1]) * ratio,
                )
            )

        resampled.append(points[-1])
        return resampled

    def compare(self, candidate: Sequence[Point], template: Sequence[Point]) -> float:
        candidate_norm = self.normalize(candidate)
        template_norm = self.normalize(template)
        if not candidate_norm or not template_norm:
            return 0.0

        total_distance = sum(
            self.distance(candidate_point, template_point)
            for candidate_point, template_point in zip(candidate_norm, template_norm)
        )
        average_distance = total_distance / min(len(candidate_norm), len(template_norm))
        return round(100.0 * max(0.0, 1.0 - average_distance / 0.75), 2)

    @staticmethod
    def path_length(stroke: Sequence[Point]) -> float:
        return sum(
            PointPatternAnalyzer.distance(previous, current)
            for previous, current in zip(stroke, stroke[1:])
        )

    @staticmethod
    def displacement(stroke: Sequence[Point]) -> Tuple[float, float]:
        if len(stroke) < 2:
            return 0.0, 0.0
        return stroke[-1][0] - stroke[0][0], stroke[-1][1] - stroke[0][1]

    @staticmethod
    def distance(point_a: Point, point_b: Point) -> float:
        return math.hypot(point_a[0] - point_b[0], point_a[1] - point_b[1])


class FixedGestureRecognizer:
    """Recognize a fixed GestureSign-style action set."""

    def __init__(
        self,
        sample_size: int = 32,
        tap_max_movement: float = 18.0,
        swipe_min_distance: float = 20.0,
        direction_ratio: float = 1.08,
    ):
        self.analyzer = PointPatternAnalyzer(sample_size=sample_size)
        self.tap_max_movement = tap_max_movement
        self.swipe_min_distance = swipe_min_distance
        self.direction_ratio = direction_ratio
        self.gestures = self._build_fixed_gestures()

    def _build_fixed_gestures(self) -> List[FixedGesture]:
        return [
            FixedGesture(
                name="media_play_pause",
                finger_count=2,
                direction=DIRECTION_TAP,
                action=GestureAction("Media Play/Pause", "2指点按", "media_play_pause"),
            ),
            FixedGesture(
                name="next_window",
                finger_count=2,
                direction=DIRECTION_RIGHT,
                action=GestureAction("Next Window", "2指右滑", "next_window"),
            ),
            FixedGesture(
                name="previous_window",
                finger_count=2,
                direction=DIRECTION_LEFT,
                action=GestureAction("Previous Window", "2指左滑", "previous_window"),
            ),
            FixedGesture(
                name="task_view",
                finger_count=3,
                direction=DIRECTION_UP,
                action=GestureAction("任务视图", "3指上滑", "task_view"),
            ),
            FixedGesture(
                name="right_click",
                finger_count=3,
                direction=DIRECTION_DOWN,
                action=GestureAction("鼠标动作", "3指下滑", "mouse_right_click"),
            ),
            FixedGesture(
                name="x1_click",
                finger_count=3,
                direction=DIRECTION_LEFT,
                action=GestureAction("鼠标动作", "3指左滑", "mouse_x1_click"),
            ),
            FixedGesture(
                name="scrcpy_middle_click_home",
                finger_count=3,
                direction=DIRECTION_RIGHT,
                action=GestureAction(
                    "鼠标动作",
                    "3指右滑",
                    "scrcpy_middle_click_home",
                ),
            ),
            FixedGesture(
                name="show_desktop",
                finger_count=4,
                direction=DIRECTION_DOWN,
                action=GestureAction(
                    "发送快捷键",
                    "4指下滑",
                    "hotkey",
                    {"keys": ["win", "d"]},
                ),
            ),
            FixedGesture(
                name="volume_down",
                finger_count=4,
                direction=DIRECTION_LEFT,
                action=GestureAction("调整音量", "4指左滑", "volume_down", {"percent": 10}),
            ),
            FixedGesture(
                name="volume_up",
                finger_count=4,
                direction=DIRECTION_RIGHT,
                action=GestureAction("鼠标动作", "4指右滑", "volume_up", {"percent": 10}),
            ),
            FixedGesture(
                name="activate_sdl_app",
                finger_count=4,
                direction=DIRECTION_UP,
                action=GestureAction(
                    "激活窗口",
                    "4指上滑",
                    "activate_window_class",
                    {
                        "class_name": "SDL_app",
                        "caption": "",
                        "is_regex": False,
                        "timeout_ms": 0,
                    },
                ),
                min_score=45.0,
            ),
        ]

    def recognize(self, strokes: Strokes, finger_count: Optional[int] = None) -> Optional[GestureMatch]:
        if finger_count is None:
            finger_count = len([stroke for stroke in strokes if stroke])
        if finger_count <= 0:
            return None

        centroid = self.analyzer.centroid_stroke(strokes)
        if not centroid:
            return None

        direction, base_score = self.classify_motion(centroid)
        if direction is None:
            return None

        candidates = [
            gesture
            for gesture in self.gestures
            if gesture.finger_count == finger_count and gesture.direction == direction
        ]
        if not candidates:
            return None

        # These fixed gestures are cardinal taps/swipes. Raw precision-touchpad
        # samples are noisy enough that template matching can reject an otherwise
        # clear directional swipe, so use the directional confidence here.
        score = base_score
        best = max(candidates, key=lambda gesture: gesture.min_score)
        if score < best.min_score:
            return None
        return GestureMatch(best, score)

    def classify_motion(self, stroke: Sequence[Point]) -> Tuple[Optional[str], float]:
        path_length = self.analyzer.path_length(stroke)
        dx, dy = self.analyzer.displacement(stroke)
        abs_dx = abs(dx)
        abs_dy = abs(dy)

        if path_length <= self.tap_max_movement:
            return DIRECTION_TAP, 100.0

        if max(abs_dx, abs_dy) < self.swipe_min_distance:
            return None, 0.0

        if abs_dx > abs_dy * self.direction_ratio:
            straightness = min(1.0, abs_dx / max(path_length, 1e-6))
            return (DIRECTION_RIGHT if dx > 0 else DIRECTION_LEFT), round(straightness * 100.0, 2)

        if abs_dy > abs_dx * self.direction_ratio:
            straightness = min(1.0, abs_dy / max(path_length, 1e-6))
            return (DIRECTION_DOWN if dy > 0 else DIRECTION_UP), round(straightness * 100.0, 2)

        return None, 0.0

    @staticmethod
    def template_for_direction(direction: str) -> Optional[Stroke]:
        templates = {
            DIRECTION_TAP: [(0.5, 0.5), (0.5, 0.5)],
            DIRECTION_UP: [(0.5, 1.0), (0.5, 0.0)],
            DIRECTION_DOWN: [(0.5, 0.0), (0.5, 1.0)],
            DIRECTION_LEFT: [(1.0, 0.5), (0.0, 0.5)],
            DIRECTION_RIGHT: [(0.0, 0.5), (1.0, 0.5)],
        }
        return templates.get(direction)


class TouchStrokeCollector:
    """Collect per-contact touch samples into multi-finger strokes.

    The collector is deliberately platform-agnostic so Qt touch events and
    Windows WM_TOUCH messages can feed the same state machine. It keeps one
    stroke per contact id and emits a completed stroke batch once the last
    active contact ends.
    """

    def __init__(self, move_threshold: float = 1.5):
        self.move_threshold = max(0.0, float(move_threshold))
        self.reset()

    def reset(self) -> None:
        self._active_strokes: Dict[int, Stroke] = {}
        self._completed_strokes: Dict[int, Stroke] = {}
        self._order: List[int] = []

    def has_active_contacts(self) -> bool:
        return bool(self._active_strokes)

    def is_active(self, contact_id: int) -> bool:
        return int(contact_id) in self._active_strokes

    def active_strokes_snapshot(self) -> Tuple[Strokes, int]:
        strokes = [list(self._active_strokes[cid]) for cid in self._order if cid in self._active_strokes]
        return strokes, len(strokes)

    def end_all_contacts(self, contact_points: Optional[Dict[int, Point]] = None) -> Optional[Tuple[Strokes, int]]:
        """End every active contact, using supplied final points when present."""
        completed: Optional[Tuple[Strokes, int]] = None
        contact_points = contact_points or {}
        for contact_id in list(self._active_strokes):
            stroke = self._active_strokes.get(contact_id) or [(0.0, 0.0)]
            point = contact_points.get(contact_id, stroke[-1])
            completed = self.process_contact(contact_id, point, "up") or completed
        return completed

    def process_contact(self, contact_id: int, point: Point, state: str) -> Optional[Tuple[Strokes, int]]:
        contact_id = int(contact_id)
        state = str(state).lower()
        point = (float(point[0]), float(point[1]))

        if state in {"cancel", "canceled", "cancelled"}:
            self.reset()
            return None

        if state == "down":
            if contact_id not in self._active_strokes:
                self._active_strokes[contact_id] = [point]
                self._order.append(contact_id)
                self._completed_strokes.pop(contact_id, None)
            else:
                self._append_point(contact_id, point)
            return None

        if state == "move":
            self._append_point(contact_id, point)
            return None

        if state in {"up", "end", "release"}:
            self._append_point(contact_id, point)
            if contact_id in self._active_strokes:
                self._completed_strokes[contact_id] = self._active_strokes.pop(contact_id)
            if not self._active_strokes and self._completed_strokes:
                strokes = [self._completed_strokes[cid] for cid in self._order if cid in self._completed_strokes]
                finger_count = len(strokes)
                self.reset()
                if strokes:
                    return strokes, finger_count
            return None

        raise ValueError(f"Unsupported touch state: {state}")

    def process_contacts(self, contacts: Iterable[Tuple[int, Point, str]]) -> Optional[Tuple[Strokes, int]]:
        completed: Optional[Tuple[Strokes, int]] = None
        for contact_id, point, state in contacts:
            completed = self.process_contact(contact_id, point, state) or completed
        return completed

    def _append_point(self, contact_id: int, point: Point) -> None:
        stroke = self._active_strokes.get(contact_id)
        if stroke is None:
            return
        if not stroke:
            stroke.append(point)
            return

        if PointPatternAnalyzer.distance(stroke[-1], point) >= self.move_threshold:
            stroke.append(point)
