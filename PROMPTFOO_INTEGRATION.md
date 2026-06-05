# Promptfoo Integration Guide

This project now includes a minimal Promptfoo integration layer that reuses the existing AItest workflow instead of replacing it.

## What Promptfoo Does Here

Promptfoo is used as the outer evaluation orchestrator:

- Generate test cases from the built-in scenario list
- Run single-scenario evaluations through the existing AItest code
- Grade results with a Python assertion
- Provide a path to later add CI/CD and red-team coverage

The existing `ai_test_script.py` remains responsible for:

- A/C dialogue execution
- D ticket generation
- B audit evaluation
- Excel-based reporting

## Added Files

- `promptfooconfig.yaml`
- `promptfoo_provider.py`
- `promptfoo_tests.py`
- `promptfoo_assert.py`

## Reused Entry Point

`ai_test_script.py` now exposes:

- `build_client()`
- `run_one_scenario_by_version()`

These functions let Promptfoo call the current scenario workflow directly by persona version.

## Quick Start

Install Python dependencies first:

```powershell
pip install -r requirements.txt
```

Run a small Promptfoo smoke test:

```powershell
npx promptfoo@latest eval -c promptfooconfig.yaml
```

Open the Promptfoo report UI:

```powershell
npx promptfoo@latest view
```

## Default Behavior

The initial config is intentionally small:

- Runs `A_v1 / B_v1 / C_v1 / D_v1`
- Uses `limit: 3`
- Keeps `skip_audit: false`
- Keeps `skip_ticket: false`

This is meant to verify the integration safely before scaling up.

## Common Next Tweaks

Run more C personas:

```yaml
persona_c_versions:
  - C_v1
  - C_v2
  - C_v3
```

Run conversation-only smoke checks:

```yaml
skip_audit: true
skip_ticket: true
```

Sample scenarios instead of using the first N:

```yaml
sample_size: 10
sample_seed: 20260605
limit: 0
```

## Suggested Rollout

1. Keep Promptfoo focused on smoke and regression orchestration first.
2. Gradually move stable business rules into deterministic Python assertions.
3. Keep B-model audit output as a second judge layer.
4. Add Promptfoo red-team coverage later for privacy, prompt injection, competitor guidance, and risky commitments.
