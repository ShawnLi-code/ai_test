import json

import ai_test_script as ai_test


def call_api(prompt, options, context):
    vars_data = context.get("vars", {})
    category = vars_data.get("category") or prompt
    if not category:
        raise ValueError("Promptfoo test vars must include `category`.")

    result = ai_test.run_one_scenario_by_version(
        category=category,
        persona_a_ver=vars_data.get("persona_a_ver", "A_v1"),
        persona_b_ver=vars_data.get("persona_b_ver", "B_v1"),
        persona_c_ver=vars_data.get("persona_c_ver", "C_v1"),
        persona_d_ver=vars_data.get("persona_d_ver") or None,
        skip_audit=bool(vars_data.get("skip_audit", False)),
        skip_ticket=bool(vars_data.get("skip_ticket", False)),
    )

    return {
        "output": json.dumps(result, ensure_ascii=False),
    }
