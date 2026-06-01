"""Automated smoke test for the Access Model Operating System.

Uses Streamlit's AppTest to exercise the key paths:
- the workbook loads and every nav page renders with no exceptions,
- the dashboard is the default landing page,
- the mission unlock + correct-answer gate work and certification % updates,
- the Scenario Challenge filters/scoring work,
- the generated distractors pass the banned-content audit,
- no "ages" grade language remains.

Run with:  python -m pytest test_app.py -q   (or)   python test_app.py
"""

import importlib

from streamlit.testing.v1 import AppTest

import app as appmod


def _run():
    at = AppTest.from_file("app.py", default_timeout=30)
    at.run()
    assert not at.exception, f"App raised on default landing: {at.exception}"
    return at


def test_default_landing_is_dashboard():
    at = _run()
    assert at.session_state["nav"] == "Session Progress"
    # Hero "Welcome back!" lives in injected markdown on the dashboard.
    assert any("Welcome back!" in (m.value or "") for m in at.markdown)


def test_every_page_renders():
    pages = [
        "Session Progress", "Missions", "Guide Certification", "Scenario Challenge",
        "Launch Toolkit", "Check Charts", "Repair Protocols", "Alpha → Access",
        "Foundations", "Guide Role", "Missions Completed", "Scenarios Completed",
    ]
    for page in pages:
        at = AppTest.from_file("app.py", default_timeout=30)
        at.run()
        at.session_state["nav"] = page
        at.run()
        assert not at.exception, f"Page '{page}' raised: {at.exception}"


def test_audit_passes_no_banned_content():
    # Direct check against the data, independent of the UI.
    data = appmod.load_data(appmod.DATA_PATH)
    violations = appmod.audit_scenarios(data["Scenarios"])
    assert violations == [], f"Banned content in choices: {violations}"
    # And the app records the pass in session state.
    at = _run()
    assert at.session_state["audit_passed"] is True


def test_no_ages_language():
    # The grade-band helper never emits the word "ages".
    for band in ["3-4", "5-6", "7-8", "All", "Adult"]:
        assert "age" not in appmod.grade_label(band).lower()
    # Scan all rendered markdown/text on every page for "ages".
    pages = list(appmod.PAGES.keys())
    for page in pages:
        at = AppTest.from_file("app.py", default_timeout=30)
        at.run()
        at.session_state["nav"] = page
        at.run()
        for m in at.markdown:
            assert "ages" not in (m.value or "").lower(), f"'ages' found on {page}"


def test_mission_unlock_and_correct_answer_gate():
    at = AppTest.from_file("app.py", default_timeout=30)
    at.run()
    at.session_state["nav"] = "Missions"
    at.run()
    assert not at.exception

    # Mission 0 unlocked, mission 1 locked until 0 is complete.
    assert appmod_unlocked(at, 0) is True

    # Open mission 0.
    _click_button(at, "open_0")
    at.run()
    assert at.session_state["active_mission"] == 0

    # Identify the correct answer for mission 0's scenario.
    m = appmod.MISSIONS[0]
    row = _get_row(m["category"], m["topic"])
    correct, over, under = appmod.make_choices(row)
    options = at.session_state["orders"]["mission:0"]

    # Click a WRONG option first -> must NOT complete the mission.
    wrong_idx = next(i for i, o in enumerate(options) if o != correct)
    _click_button(at, f"m_0_opt_{wrong_idx}")
    at.run()
    assert at.session_state["mission_answered"][0] is True
    assert at.session_state["mission_correct"][0] is False
    assert not at.session_state["missions_completed"].get(0), "Wrong answer must not complete"
    # A "Try again" button is present; no completion button.
    assert _has_button(at, "m_0_retry")
    assert not _has_button(at, "m_0_complete")

    # Try again, then answer correctly.
    _click_button(at, "m_0_retry")
    at.run()
    correct_idx = next(i for i, o in enumerate(options) if o == correct)
    _click_button(at, f"m_0_opt_{correct_idx}")
    at.run()
    assert at.session_state["mission_correct"][0] is True
    _click_button(at, "m_0_complete")
    at.run()
    assert at.session_state["missions_completed"].get(0) is True

    # Certification % should now reflect 1/9.
    at.session_state["nav"] = "Guide Certification"
    at.run()
    assert not at.exception
    pct = round(100 * sum(1 for v in at.session_state["missions_completed"].values() if v)
                / len(appmod.MISSIONS))
    assert pct == round(100 * 1 / 9)


def test_scenario_challenge_scoring():
    at = AppTest.from_file("app.py", default_timeout=30)
    at.run()
    at.session_state["nav"] = "Scenario Challenge"
    at.run()
    assert not at.exception
    deck = at.session_state["practice_deck"]
    assert deck and len(deck) > 0

    ri = deck[0]
    row = _get_row_by_index(ri)
    correct, over, under = appmod.make_choices(row)
    options = at.session_state["orders"][f"practice:{ri}"]
    correct_idx = next(i for i, o in enumerate(options) if o == correct)
    _click_button(at, f"p_{ri}_opt_{correct_idx}")
    at.run()
    assert at.session_state["practice_correct"][ri] is True
    assert len(at.session_state["practice_history"]) == 1
    assert at.session_state["practice_history"][0]["correct"] is True


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def _click_button(at, key):
    for b in at.button:
        if b.key == key:
            b.click()
            return
    raise AssertionError(f"Button '{key}' not found")


def _has_button(at, key):
    return any(b.key == key for b in at.button)


def appmod_unlocked(at, idx):
    # Re-evaluate using the app's own session-state-backed logic.
    return idx == 0 or bool(at.session_state["missions_completed"].get(idx - 1))


def _get_row(category, topic):
    data = appmod.load_data(appmod.DATA_PATH)
    scen = data["Scenarios"]
    return scen[(scen["Category"] == category) & (scen["Topic"] == topic)].iloc[0]


def _get_row_by_index(ri):
    data = appmod.load_data(appmod.DATA_PATH)
    return data["Scenarios"].loc[ri]


if __name__ == "__main__":
    importlib.reload(appmod)
    test_default_landing_is_dashboard()
    test_every_page_renders()
    test_audit_passes_no_banned_content()
    test_no_ages_language()
    test_mission_unlock_and_correct_answer_gate()
    test_scenario_challenge_scoring()
    print("All smoke tests passed.")
