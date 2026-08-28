# Security Spec for Data Separation

## Data Invariants
- A student, attendance session, follow-up, or monthly recap MUST belong to a user, identified by `userId`.
- Access to these documents MUST be restricted to the owner (`userId == request.auth.uid`).

## The "Dirty Dozen" Payloads (Simplified)
1. Read student belonging to another user.
2. Write student with `userId` not matching `request.auth.uid`.
3. Create attendance session with `userId` not matching `request.auth.uid`.
4. Update attendance session belonging to another user.
...
