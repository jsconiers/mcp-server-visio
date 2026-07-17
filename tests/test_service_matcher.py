"""
Tests for Azure service-name resolution.

Runs on ANY platform. visio_client imports win32com/pythoncom at module level,
which is why this repo had no tests -- it cannot be imported off Windows. But
the service matcher is pure Python and has nothing to do with COM, so stubbing
those two modules makes it testable on macOS, Linux, and in CI.

Nothing here touches Visio. The COM paths still need a Windows box.

Run: python tests/test_service_matcher.py
"""
import os
import sys
import types

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# Stub the Windows-only COM modules before importing visio_client.
for _m in ("win32com", "win32com.client", "pythoncom"):
    sys.modules.setdefault(_m, types.ModuleType(_m))
sys.modules["win32com"].client = sys.modules["win32com.client"]

from visio_client import VisioClient, _fuzzy_key, _norm  # noqa: E402


def _client():
    """VisioClient without touching COM -- __init__ only loads the JSON map."""
    return VisioClient()


def test_stencil_map_loads():
    c = _client()
    assert len(c._stencil_map) == 209, f"expected 209, got {len(c._stencil_map)}"
    assert "azure/front-door" in c._stencil_map
    print("  PASS: test_stencil_map_loads")


def test_fuzzy_key_still_builds_keys():
    """_fuzzy_key CONSTRUCTS a key; _norm normalizes for COMPARISON. Distinct jobs."""
    assert _fuzzy_key("App Services") == "azure/app-services"
    assert _fuzzy_key("front_door") == "azure/front-door"
    assert _fuzzy_key("azure/sql-database") == "azure/sql-database"
    assert _norm("Front Door And Cdn Profiles") == "frontdoorandcdnprofiles"
    print("  PASS: test_fuzzy_key_still_builds_keys")


def test_resolution_paths():
    c = _client()
    assert c._resolve_azure_service("azure/front-door")[1] == "Front Door And Cdn Profiles"
    assert c._resolve_azure_service("front door") is not None      # via _fuzzy_key
    assert c._resolve_azure_service("App Services") is not None    # via _fuzzy_key
    # display-name match -- this is new; _fuzzy_key alone cannot do it
    assert c._resolve_azure_service("Kubernetes Services") is not None
    print("  PASS: test_resolution_paths")


def test_suggestions_catch_the_substring_blind_spot():
    """
    The regression this fix exists for.

    'azure/functions' is not a substring of 'azure/function-apps', so pure
    containment returns nothing and the user is told only "unknown".
    """
    c = _client()
    assert c._resolve_azure_service("azure/functions") is None, "should not resolve"
    s = c.suggest_services("azure/functions")
    assert "azure/function-apps" in s, f"expected function-apps in {s}"

    assert "azure/kubernetes-services" in c.suggest_services("azure/kubernetes")
    assert "azure/cosmos-db" in c.suggest_services("cosmos")
    print("  PASS: test_suggestions_catch_the_substring_blind_spot")


def test_nonsense_gets_no_invented_suggestion():
    """Better to say nothing than to point at a plausible wrong service."""
    c = _client()
    assert c.suggest_services("qwertyuiop-not-a-service") == []
    assert c.suggest_services("") == []
    print("  PASS: test_nonsense_gets_no_invented_suggestion")


def test_suggestions_are_real_keys_and_bounded():
    c = _client()
    for probe in ("azure/functions", "storage", "sql", "gateway", "cosmos"):
        s = c.suggest_services(probe)
        assert len(s) <= 5, f"{probe}: {len(s)} suggestions, cap is 5"
        assert len(s) == len(set(s)), f"{probe}: duplicates in {s}"
        for k in s:
            assert k in c._stencil_map, f"{probe}: suggested non-existent key {k}"
    print("  PASS: test_suggestions_are_real_keys_and_bounded")


if __name__ == "__main__":
    print("Running service matcher tests...\n")
    test_stencil_map_loads()
    test_fuzzy_key_still_builds_keys()
    test_resolution_paths()
    test_suggestions_catch_the_substring_blind_spot()
    test_nonsense_gets_no_invented_suggestion()
    test_suggestions_are_real_keys_and_bounded()
    print("\nAll tests passed!")
