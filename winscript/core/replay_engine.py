import time
from winscript.core.workflow import Workflow, WORKFLOW_DIR

# Tool registry — maps tool_name string → actual function
# Populated at server startup
_tool_registry: dict = {}

def register_tool(name: str, fn):
    _tool_registry[name] = fn

def replay_workflow(name: str, dry_run: bool = False) -> str:
    """
    Replay a saved workflow step by step.
    dry_run=True: shows what would be executed without running it.
    """
    wf = Workflow.load(name)
    if wf is None:
        return f"ERROR: Workflow '{name}' not found."

    lines = [f"Replaying '{wf.name}' ({len(wf.steps)} steps)"]
    if dry_run:
        lines.append("[DRY RUN — no actions will execute]")

    for step in wf.steps:
        fn = _tool_registry.get(step.tool_name)
        if fn is None:
            lines.append(f"  Step {step.step_number}: SKIP — tool '{step.tool_name}' not in registry")
            continue

        if dry_run:
            lines.append(f"  Step {step.step_number}: {step.tool_name}({step.args})")
            continue

        start = time.time()
        try:
            result = fn(**step.args)
            duration = (time.time() - start) * 1000
            lines.append(f"  Step {step.step_number} ✓ {step.tool_name} → {str(result)[:60]} [{duration:.0f}ms]")
        except Exception as e:
            lines.append(f"  Step {step.step_number} ✗ {step.tool_name} → ERROR: {str(e)}")

    # Update run count
    if not dry_run:
        wf.run_count += 1
        wf.last_run_at = time.time()
        wf.save()

    return "\n".join(lines)