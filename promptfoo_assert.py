import json


def get_assert(output, context):
    try:
        data = json.loads(output)
    except json.JSONDecodeError as exc:
        return {
            "pass": False,
            "score": 0,
            "reason": f"Result is not valid JSON: {exc}",
        }

    error = data.get("error")
    passed = data.get("passed")
    violation_count = data.get("violation_count", 0)
    reason = data.get("violation_summary") or "Scenario passed"

    if error:
        return {
            "pass": False,
            "score": 0,
            "reason": error,
            "namedScores": {
                "violations": float(violation_count or 0),
            },
        }

    if passed is None:
        return {
            "pass": True,
            "score": 1,
            "reason": "Audit skipped by configuration",
            "namedScores": {
                "violations": 0.0,
            },
        }

    score = 1.0 if passed else 0.0
    return {
        "pass": bool(passed),
        "score": score,
        "reason": reason,
        "namedScores": {
            "violations": float(violation_count or 0),
        },
    }
