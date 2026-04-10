from __future__ import annotations

SEED_WORKSPACE_FIXTURES = [
    {
        "id": "ws_ops",
        "name": "Operations Hub",
        "slug": "operations-hub",
        "description": "Execution boards, handoffs, and launch tracking.",
        "color": "#0f766e",
        "position": 0,
        "member_id": "wm_ops_owner",
        "workbook": {
            "id": "wb_ops",
            "name": "Ops workbook",
            "description": "Primary workbook for execution planning and launches.",
            "position": 0,
            "sheets": [
                {
                    "id": "sheet_ops_launch",
                    "name": "Launch tracker",
                    "key": "launch_tracker",
                    "position": 0,
                    "rows": [
                        {
                            "id": "row_1",
                            "task": "Finalize launch checklist",
                            "owner": "Marta",
                            "status": "In review",
                            "timeline": "Apr 11",
                            "progress": 86,
                        },
                        {
                            "id": "row_2",
                            "task": "QA smoke for billing",
                            "owner": "Ivan",
                            "status": "Blocked",
                            "timeline": "Apr 12",
                            "progress": 42,
                        },
                        {
                            "id": "row_3",
                            "task": "Sync support macros",
                            "owner": "Sara",
                            "status": "Ready",
                            "timeline": "Apr 13",
                            "progress": 100,
                        },
                    ],
                },
                {
                    "id": "sheet_ops_vendor",
                    "name": "Vendor intake",
                    "key": "vendor_intake",
                    "position": 1,
                    "rows": [
                        {
                            "id": "row_4",
                            "task": "Security review",
                            "owner": "Nick",
                            "status": "Pending",
                            "timeline": "Apr 18",
                            "progress": 24,
                        },
                        {
                            "id": "row_5",
                            "task": "Contract redlines",
                            "owner": "Julia",
                            "status": "In progress",
                            "timeline": "Apr 19",
                            "progress": 63,
                        },
                    ],
                },
            ],
        },
    },
    {
        "id": "ws_strategy",
        "name": "Strategy Studio",
        "slug": "strategy-studio",
        "description": "Quarter planning, themes, and executive rollups.",
        "color": "#1d4ed8",
        "position": 1,
        "member_id": "wm_strategy_owner",
        "workbook": {
            "id": "wb_strategy",
            "name": "Strategy workbook",
            "description": "Primary workbook for roadmap and planning documents.",
            "position": 0,
            "sheets": [
                {
                    "id": "sheet_strategy_roadmap",
                    "name": "Roadmap",
                    "key": "roadmap",
                    "position": 0,
                    "rows": [
                        {
                            "id": "row_6",
                            "task": "North star metrics",
                            "owner": "Lina",
                            "status": "Ready",
                            "timeline": "Apr 16",
                            "progress": 71,
                        },
                        {
                            "id": "row_7",
                            "task": "Pricing hypotheses",
                            "owner": "Oleg",
                            "status": "Draft",
                            "timeline": "Apr 22",
                            "progress": 31,
                        },
                    ],
                }
            ],
        },
    },
]