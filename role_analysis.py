ROLE_FEATURES = {
    # ----------------------------
    # FORWARDS (C / LW / RW)
    # ----------------------------
    "F": {
        # --- Core line roles ---
        "Scorer": {
            "minutes_basis": "toi_ev",   # shrinkage minutes
            "k": 600,
            "features": {
                "ixg60_ev": 0.40,
                "shots60_ev": 0.25,
                "g60_ev": 0.15,
                "ixg60_pp": 0.10,
                "g60_pp": 0.10,
            },
            "notes": "Finishing pressure + shot quality, with small PP influence."
        },
        "Distributor": {
            "minutes_basis": "toi_ev",
            "k": 600,
            "features": {
                "a1_60_ev": 0.45,
                "p1_60_ev": 0.20,
                "a1_60_pp": 0.20,
                "on_ice_xgf60_ev": 0.15,
            },
            "notes": "Primary creation + involvement; on-ice xGF as shot-generation proxy."
        },
        "Power": {
            "minutes_basis": "toi_ev",
            "k": 600,
            "features": {
                "hd_attempts60_ev": 0.30,     # if not available, swap -> ixg60_ev
                "pen_drawn60": 0.20,
                "hits60": 0.15,
                "ixg60_ev": 0.15,
                "weight_lb": 0.10,
                "g60_ev": 0.10,
            },
            "notes": "Net-front pressure + puck protection proxies + body."
        },
        "Forecheck": {
            "minutes_basis": "toi_ev",
            "k": 600,
            "features": {
                "hits60": 0.25,
                "takeaways60": 0.20,
                "pen_drawn60": 0.15,
                "on_ice_xgf60_ev": 0.20,
                "p1_60_ev": 0.10,
                "blocks60": 0.10,
            },
            "notes": "Pressure/retrieval; on-ice offense proxy; keep blocks modest."
        },
        "Shutdown": {
            "minutes_basis": "toi_pk",
            "k": 150,
            "features": {
                "share_pk": 0.35,
                "toi_pk": 0.20,
                "on_ice_xga60_ev": -0.25,     # negative means lower is better
                "blocks60": 0.10,
                "pen_diff60": 0.10,
            },
            "notes": "PK usage + suppression proxy; xGA uses EV context."
        },

        # --- Special teams roles (optional separate) ---
        "PP_Specialist": {
            "minutes_basis": "toi_pp",
            "k": 150,
            "features": {
                "share_pp": 0.35,
                "p1_60_pp": 0.25,
                "a1_60_pp": 0.20,
                "ixg60_pp": 0.10,
                "shots60_pp": 0.10,
            },
            "notes": "PP driver blend; balances creation and threat."
        },
        "PK_Specialist": {
            "minutes_basis": "toi_pk",
            "k": 150,
            "features": {
                "share_pk": 0.40,
                "toi_pk": 0.20,
                "blocks60": 0.20,
                "on_ice_xga60_ev": -0.20,
            },
            "notes": "Usage + shot suppression proxies."
        },

        # --- Utility / discipline (optional) ---
        "Discipline": {
            "minutes_basis": "toi_ev",
            "k": 600,
            "features": {
                "pen_diff60": 0.60,
                "pen_taken60": -0.20,
                "giveaways60": -0.20,
            },
            "notes": "Helps identify 'doesn't hurt you' profiles."
        },
    },

    # ----------------------------
    # DEFENSE (LD / RD)
    # ----------------------------
    "D": {
        "PuckMover": {
            "minutes_basis": "toi_ev",
            "k": 700,
            "features": {
                "a1_60_ev": 0.25,
                "p1_60_ev": 0.15,
                "on_ice_xgf60_ev": 0.35,
                "giveaways60": -0.15,
                "takeaways60": 0.10,
            },
            "notes": "Transition/activation proxy; on-ice xGF stands in for breakouts/entries."
        },
        "QB_PP": {
            "minutes_basis": "toi_pp",
            "k": 200,
            "features": {
                "share_pp": 0.35,
                "a1_60_pp": 0.30,
                "p1_60_pp": 0.20,
                "shots60_pp": 0.10,
                "ixg60_pp": 0.05,
            },
            "notes": "PP quarterback; heavy emphasis on assists/primary points."
        },
        "ShutdownD": {
            "minutes_basis": "toi_pk",
            "k": 200,
            "features": {
                "share_pk": 0.25,
                "toi_pk": 0.20,
                "on_ice_xga60_ev": -0.25,
                "blocks60": 0.20,
                "hits60": 0.10,
            },
            "notes": "Suppression + PK usage + blocks; keep hits modest."
        },
        "PhysicalD": {
            "minutes_basis": "toi_ev",
            "k": 700,
            "features": {
                "hits60": 0.30,
                "blocks60": 0.20,
                "weight_lb": 0.20,
                "pen_taken60": -0.15,
                "share_pk": 0.15,
            },
            "notes": "Net-front clearing proxy; discipline matters."
        },
        "TwoWayD": {
            "minutes_basis": "toi_ev",
            "k": 700,
            # This role is computed from other role scores; leave features empty for now.
            "features": {},
            "notes": "Compute as balanced high PuckMover + ShutdownD with low Giveaways."
        },
    },

    # ----------------------------
    # FEATURE ALIASES / FALLBACKS
    # ----------------------------
    "FALLBACKS": {
        # If you don't have hd_attempts60_ev, use ixg60_ev as proxy for net-front pressure.
        "hd_attempts60_ev": ["hd_attempts60_ev", "ixg60_ev"],
        # If you don't have on-ice xG, allow CF/FF proxies (fill with what you have).
        "on_ice_xgf60_ev": ["on_ice_xgf60_ev", "on_ice_cf60_ev", "on_ice_ff60_ev"],
        "on_ice_xga60_ev": ["on_ice_xga60_ev", "on_ice_ca60_ev", "on_ice_fa60_ev"],
    },

    # ----------------------------
    # POSITION RULES (your choice)
    # ----------------------------
    "POSITION_RULES": {
        # Assumes you have a coarse position flag and handedness/shoots.
        # You can override with explicit roster position if available later.
        "derive_position": {
            "D": {"L": "LD", "R": "RD"},
            "W": {"L": "LW", "R": "RW"},
            "C": {"L": "C",  "R": "C"},
        }
    },
}