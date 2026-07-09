# Logout

**Ribbon:** `btnLogout` — “Logout”  
**Login required:** No — enabled when a **target instance** exists (committed auto-login prefs or live `InstanceUrl`)  
**COM / VBA:** None

See also: [common-performance-requirements.md](./common-performance-requirements.md)

---

## Summary

End the Salesforce session and clear the committed target instance so the next gated API call shows login options (org/sandbox switch). Persisted refresh tokens remain in the credential store so the same host can silent-refresh after re-selection.

---

## Functional requirements

### FR-LO-1 Trigger

- Ribbon **Logout** enabled only when a target instance exists (committed prefs with `ShowLoginOptionsOnNextUse == false`, or live `InstanceUrl`).
- Click is not wrapped in the login gate; when disabled, the button is unavailable.

### FR-LO-2 Session teardown

On Logout:

1. Invalidate `SessionContext` — clear access token, in-memory refresh token, `InstanceUrl`, identity URL, display name, API version.
2. Save login preferences with the same tenant/environment/sandbox/API caches and `ShowLoginOptionsOnNextUse = true` (clears the **target** so the next gated call shows login options).
3. **Do not** delete persisted refresh tokens from the credential store (tokens remain keyed by host).
4. Clear in-memory host-key / display-name caches used by the gate.

### FR-LO-3 User feedback

- Acknowledge when Logout cleared a live session and/or committed target (modal or equivalent).
- When already cleared (no target, no session) → no-op or brief info (FR-LO-6).

### FR-LO-4 UI state

- Group label: `Force.com Connector Next Generation (no logon user)`.
- Disable Logout; Options stays enabled (global connector options).
- Invalidate ribbon controls after Logout so enablement and label refresh immediately.

### FR-LO-5 No Excel data mutation

- Logout does not modify workbook tables or selections.

### FR-LO-6 Idempotent

- Logout when already logged out with no target → no-op or brief info (no error).

---

## Non-functional requirements

### NFR-LO-1 Session clear (no SOAP)

Local session discard is sufficient; remote access tokens may remain valid until expiry. Refresh tokens stay in the OS credential store until overwritten by a later sign-in for that host. No SOAP/WCF logout call.

### NFR-LO-2 Security

- Access tokens and in-memory secrets cleared on Logout.
- No token values in logs.
- Persisted refresh tokens remain for convenience; they are host-scoped and only used after the user re-commits a target via login options.

### NFR-LO-3 Performance

- Synchronous local clear; must not block Excel indefinitely.

### NFR-LO-4 Testability

- Core session clear logic unit tested without UI (`SessionContext`, `SessionGate`, credential store fakes).

---

## Intentionally dropped

| Topic | Reason |
|-------|--------|
| SOAP logout call | SOAP stack not used |
| METAAPI client teardown | Translation Helper out of scope |
| Deleting credential-store refresh token on Logout | User can switch orgs without re-authorizing the previous host |

---

## Related

- Login flows via `SessionGate` / `SessionLoginOrchestrator`.
- Successful login refreshes the group label (instance + display name) and enables Logout.
