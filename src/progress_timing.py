"""Stage-based ETA estimates and a small history of successful desktop runs."""

from __future__ import annotations

import json
import math
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from statistics import median
from typing import Callable, Mapping, Sequence


HISTORY_LIMIT = 5


@dataclass(frozen=True)
class TimeEstimate:
    status: str
    remaining_seconds: float | None = None

    @property
    def text(self) -> str:
        if self.status == "complete":
            return "已完成"
        if self.status == "stopped":
            return "已停止"
        if self.status == "reestimating":
            return "預估剩餘：重新估算中"
        if self.remaining_seconds is None:
            return "預估剩餘：估算中"
        seconds = max(1, math.ceil(self.remaining_seconds))
        minutes, seconds = divmod(seconds, 60)
        hours, minutes = divmod(minutes, 60)
        duration = f"{hours}:{minutes:02}:{seconds:02}" if hours else f"{minutes:02}:{seconds:02}"
        return f"預估剩餘：{duration}"


class TimingHistory:
    def __init__(self, path: Path, stages: Sequence[str]) -> None:
        self.path = path
        self.stages = tuple(stages)

    def _valid_sample(self, sample) -> bool:
        return (
            isinstance(sample, dict)
            and set(sample) == set(self.stages)
            and all(
                isinstance(value, (int, float)) and not isinstance(value, bool)
                and math.isfinite(value) and value >= 0
                for value in sample.values()
            )
        )

    def _read(self) -> dict[str, list[dict[str, float]]]:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError, UnicodeError):
            return {}
        if not isinstance(data, dict) or data.get("version") != 1 or not isinstance(data.get("customers"), dict):
            return {}
        customers = {}
        for customer, samples in data["customers"].items():
            if customer in ("AVTC", "RAKEN") and isinstance(samples, list):
                customers[customer] = [sample for sample in samples if self._valid_sample(sample)][-HISTORY_LIMIT:]
        return customers

    def baseline(self, customer: str) -> dict[str, float]:
        samples = self._read().get(customer, [])
        if not samples:
            return {}
        return {stage: median(sample[stage] for sample in samples) for stage in self.stages}

    def record(self, customer: str, durations: Mapping[str, float]) -> None:
        sample = dict(durations)
        if customer not in ("AVTC", "RAKEN") or not self._valid_sample(sample):
            return
        customers = self._read()
        customers[customer] = [*customers.get(customer, []), sample][-HISTORY_LIMIT:]
        data = json.dumps({"version": 1, "customers": customers}, ensure_ascii=False, indent=2)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=self.path.parent, prefix="timings-", suffix=".tmp", delete=False) as handle:
                temporary = Path(handle.name)
                handle.write(data)
            temporary.replace(self.path)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)


class RemainingTimeEstimator:
    def __init__(
        self,
        stages: Sequence[str],
        history: Mapping[str, float] | None = None,
        *,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self.stages = tuple(stages)
        self.history = dict(history or {})
        self.clock = clock
        self.active_stage: str | None = None
        self.durations: dict[str, float] = {}
        self.observed_at = clock()
        self.active_elapsed_at_observation = 0.0
        self.status: str | None = None

    def observe(self, label: str, elapsed_seconds: float, durations: Sequence[tuple[str, float]]) -> None:
        self.durations = dict(durations)
        self.active_stage = label
        self.observed_at = self.clock()
        self.active_elapsed_at_observation = max(0.0, elapsed_seconds - sum(self.durations.values()))
        if label == "Done":
            self.stop(success=True)

    def stop(self, *, success: bool) -> None:
        self.status = "complete" if success else "stopped"

    def _expected_durations(self) -> dict[str, float]:
        if set(self.history) == set(self.stages) and sum(self.history.values()) >= 1:
            # Compare completed, substantial phases to adjust for this run's size/speed.
            comparable = [stage for stage in self.durations if self.history.get(stage, 0) >= 0.5]
            historical = sum(self.history[stage] for stage in comparable)
            observed = sum(self.durations[stage] for stage in comparable)
            scale = max(0.25, min(4.0, observed / historical)) if historical and observed >= 0.2 else 1.0
            return {stage: seconds * scale for stage, seconds in self.history.items()}
        # Fast discovery phases are poor predictors of Excel parsing/writing.
        substantial = [seconds for seconds in self.durations.values() if seconds >= 0.5]
        if len(self.durations) < 4 or len(substantial) < 2:
            return {}
        typical = median(substantial)
        return {stage: typical for stage in self.stages}

    def snapshot(self) -> TimeEstimate:
        if self.status is not None:
            return TimeEstimate(self.status)
        if self.active_stage not in self.stages:
            return TimeEstimate("estimating")
        expected = self._expected_durations()
        if not expected:
            return TimeEstimate("estimating")
        active_elapsed = self.active_elapsed_at_observation + max(0.0, self.clock() - self.observed_at)
        active_remaining = expected[self.active_stage] - active_elapsed
        if active_remaining <= 0:
            return TimeEstimate("reestimating")
        index = self.stages.index(self.active_stage)
        remaining = active_remaining + sum(expected[stage] for stage in self.stages[index + 1:])
        return TimeEstimate("running", remaining)
