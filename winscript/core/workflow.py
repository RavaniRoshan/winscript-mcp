import json
import time
import uuid
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Callable, Optional

WORKFLOW_DIR = Path.home() / ".winscript" / "workflows"
WORKFLOW_DIR.mkdir(parents=True, exist_ok=True)

@dataclass
class WorkflowStep:
    step_number: int
    tool_name: str
    args: dict
    result: str
    duration_ms: float
    state_delta: dict = field(default_factory=dict)
    timestamp: float = field(default_factory=time.time)

@dataclass
class Workflow:
    name: str
    description: str
    steps: list[WorkflowStep]
    created_at: float = field(default_factory=time.time)
    run_count: int = 0
    last_run_at: Optional[float] = None
    workflow_id: str = field(default_factory=lambda: str(uuid.uuid4())[:8])

    def save(self):
        path = WORKFLOW_DIR / f"{self.name}.json"
        path.write_text(json.dumps(asdict(self), indent=2))
        return str(path)

    @classmethod
    def load(cls, name: str) -> Optional["Workflow"]:
        path = WORKFLOW_DIR / f"{name}.json"
        if not path.exists():
            return None
        data = json.loads(path.read_text())
        steps = [WorkflowStep(**s) for s in data.pop("steps")]
        return cls(steps=steps, **data)

    @classmethod
    def list_all(cls) -> list[dict]:
        results = []
        for f in WORKFLOW_DIR.glob("*.json"):
            try:
                data = json.loads(f.read_text())
                results.append({
                    "name": data.get("name", f.stem),
                    "description": data.get("description", ""),
                    "steps": len(data.get("steps", [])),
                    "run_count": data.get("run_count", 0),
                    "created_at": data.get("created_at", 0),
                })
            except Exception:
                continue
        return results

class WorkflowRecorder:
    """Records tool calls into a replayable workflow."""

    def __init__(self):
        self._recording: bool = False
        self._steps: list[WorkflowStep] = []
        self._name: str = ""
        self._description: str = ""

    @property
    def is_recording(self) -> bool:
        return self._recording

    def start(self, name: str, description: str = "") -> str:
        if self._recording:
            return f"ERROR: Already recording '{self._name}'. Stop it first."
        self._recording = True
        self._steps = []
        self._name = name
        self._description = description
        return f"Recording started: '{name}'. Every tool call will be captured."

    def record_step(self, tool_name: str, args: dict, result: str,
                    duration_ms: float, state_delta: dict = None):
        if not self._recording:
            return
        self._steps.append(WorkflowStep(
            step_number=len(self._steps) + 1,
            tool_name=tool_name,
            args=args,
            result=result,
            duration_ms=duration_ms,
            state_delta=state_delta or {},
        ))

    def stop(self) -> str:
        if not self._recording:
            return "ERROR: Not currently recording."
        if not self._steps:
            self._recording = False
            return "Recording stopped. No steps captured."
        wf = Workflow(
            name=self._name,
            description=self._description,
            steps=self._steps,
        )
        path = wf.save()
        self._recording = False
        count = len(self._steps)
        self._steps = []
        return f"Workflow '{self._name}' saved: {count} steps → {path}"

    def discard(self) -> str:
        self._recording = False
        self._steps = []
        return "Recording discarded."

# Singleton
recorder = WorkflowRecorder()