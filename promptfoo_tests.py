import ai_test_script as ai_test


def create_tests(config=None):
    config = config or {}
    scenarios = list(ai_test.DEFAULT_SCENARIOS)

    limit = int(config.get("limit", 0) or 0)
    sample_size = int(config.get("sample_size", 0) or 0)
    sample_seed = int(config.get("sample_seed", 42) or 42)
    selected = ai_test.select_scenarios(
        scenarios,
        limit=limit,
        sample_size=sample_size,
        sample_seed=sample_seed,
    )

    persona_a_ver = config.get("persona_a_ver", "A_v1")
    persona_b_ver = config.get("persona_b_ver", "B_v1")
    persona_d_ver = config.get("persona_d_ver", "")
    persona_c_versions = config.get("persona_c_versions", ["C_v1"])
    skip_audit = bool(config.get("skip_audit", False))
    skip_ticket = bool(config.get("skip_ticket", False))

    tests = []
    for persona_c_ver in persona_c_versions:
        for category in selected:
            tests.append(
                {
                    "description": f"{category} | {persona_a_ver}/{persona_b_ver}/{persona_c_ver}",
                    "vars": {
                        "category": category,
                        "persona_a_ver": persona_a_ver,
                        "persona_b_ver": persona_b_ver,
                        "persona_c_ver": persona_c_ver,
                        "persona_d_ver": persona_d_ver,
                        "skip_audit": skip_audit,
                        "skip_ticket": skip_ticket,
                    },
                }
            )
    return tests
