# Security Specification & Test Payloads (Phase 0)

## 1. Data Invariants
- Admin status cannot be forged or self-assigned in `users` collection.
- Admin status is verified via `exists(/databases/$(database)/documents/admins/$(request.auth.uid))` or bootstrapped admin email `robelmd2124@gmail.com` with `email_verified == true`.
- Campaigns require valid string IDs, non-empty page names, strictly positive numbers for budget and rate, and immutable recordId.
- Settings collection is read-only for authenticated viewers, and write-protected strictly for administrators.
- Payments require valid recordId and non-negative paid amount.

## 2. The "Dirty Dozen" Payloads (Must be rejected)
1. **Unauthenticated Read on Admins**: `{ "path": "/admins/admin1", "auth": null }` -> Denied
2. **Unauthenticated Write on Campaigns**: `{ "path": "/campaigns/camp1", "data": { "pageName": "Test" }, "auth": null }` -> Denied
3. **Ghost / Shadow Field Injection in Campaigns**: `{ "path": "/campaigns/camp1", "data": { "recordId": "camp1", "pageName": "Test", "budget": 100, "month": "October", "year": "2026", "isSuperAdmin": true } }` -> Denied by strict schema checks
4. **Negative Budget Campaign**: `{ "path": "/campaigns/camp2", "data": { "recordId": "camp2", "pageName": "Test", "budget": -50, "month": "October", "year": "2026" } }` -> Denied
5. **Denial of Wallet Long ID**: `{ "path": "/campaigns/" + "a".repeat(200), "data": { ... } }` -> Denied by isValidId length <= 128
6. **Non-admin Overwrite on Settings**: `{ "path": "/settings/global", "data": { "dollarRate": 150 }, "auth": { "uid": "viewer_1", "email": "viewer@example.com" } }` -> Denied
7. **Email Spoofing (Unverified Email)**: `{ "auth": { "uid": "fake_admin", "email": "robelmd2124@gmail.com", "email_verified": false } }` -> Denied
8. **Malicious Empty Page Name**: `{ "path": "/campaigns/camp3", "data": { "recordId": "camp3", "pageName": "", "budget": 50, "month": "October", "year": "2026" } }` -> Denied
9. **Tampered Immutable Record ID during update**: `{ "path": "/campaigns/camp1", "data": { "recordId": "tampered_id" } }` -> Denied
10. **Arbitrary Write to User Profile by another user**: `{ "path": "/users/victim_user", "data": { "role": "admin" }, "auth": { "uid": "attacker_user" } }` -> Denied
11. **Self-Escalation to Admin in User Document**: `{ "path": "/users/attacker_user", "data": { "role": "admin" } }` -> Denied
12. **Blind Collection Query / List Dump**: `{ "path": "/admins", "auth": null }` -> Denied
