---
schema_version: 1
intent_version: 1
mode: plan-only
objective: to_confirm
requested_artifacts: []
stop_condition: to_confirm
input_scope: []
official_constraints: []
logic_chain: []
unresolved_items: []
execution_requested: false
---

# Workflow Intent

Use this structure for complex work or when the intent needs to be archived.

- `plan-only`: stop at the requested analysis, template, gap report, or review. Keep `execution_requested: false`.
- `execute`: use only when the user requested file mutation or command execution. Set `execution_requested: true`, then detect the platform and select a checked-in adapter.
- `replan`: increment `intent_version` and review affected planning when the objective, input scope, official constraints, requested artifacts, stop condition, or logic chain changes.
- `reroute`: keep the semantic intent and `intent_version` unchanged when only the operating system, Shell, or tool availability changes; select a different adapter only if execution is requested.

Small, single-step requests may keep the same structure in the conversation without writing an intent file.
