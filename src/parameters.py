"""Parameter loading and validation for the Space RDS pipeline."""
from __future__ import annotations
import datetime as dt
from dataclasses import dataclass, field
import yaml


def _d(s):
    if isinstance(s, dt.date):
        return s if not isinstance(s, dt.datetime) else s.date()
    return dt.date.fromisoformat(str(s))


@dataclass
class Params:
    quarter: str
    as_at: dt.date
    rpf_bands: list                  # [(months, factor)]
    leo_debris_bands: list           # [(months, factor)]
    scenarios: dict
    equity_by_year: dict
    s3123_factors: list              # [(from_date, factor)]
    s3123_qs: dict
    igr_qs: dict
    igr_xol: dict
    outwards_slots: list
    entities: list
    ingest: dict
    raw: dict = field(repr=False, default_factory=dict)
    # S2126 consortium sub-share schedule [(from_date, factor)] — the new
    # consortium participant from 2026-04-01 (5/30). Empty if the quarter's
    # config predates S2126, in which case s2126_factor() returns 0 throughout.
    s2126_factors: list = field(default_factory=list)

    @classmethod
    def load(cls, path: str) -> "Params":
        # Always UTF-8: the config carries em-dashes / arrows in comments, and
        # Windows would otherwise default to cp1252 and crash on those bytes.
        with open(path, encoding="utf-8") as f:
            c = yaml.safe_load(f)
        return cls(
            quarter=c["quarter"],
            as_at=_d(c["as_at_date"]),
            rpf_bands=[(b["months"], b["factor"]) for b in c["rpf_bands"]],
            leo_debris_bands=[(b["months"], b["factor"]) for b in c["leo_debris_rpf_bands"]],
            scenarios=c["scenarios"],
            equity_by_year={int(k): v for k, v in c["equity_share_by_year"].items()},
            s3123_factors=sorted(
                [(_d(x["from"]), x["factor"]) for x in c["s3123_consortium_factors"]]
            ),
            s3123_qs={
                "cession": c["s3123_qs"]["cession"],
                "date_from": _d(c["s3123_qs"]["date_from"]),
                "date_to": _d(c["s3123_qs"]["date_to"]),
                "excluded": set(c["s3123_qs"]["excluded_spacecraft"]),
            },
            igr_qs=c["igr_qs"],
            igr_xol=c["igr_xol"],
            outwards_slots=[
                {**s, "from": _d(s["from"]), "to": _d(s["to"])}
                for s in c["outwards_ri_slots"]
            ],
            entities=c["entities"],
            ingest=c["ingest"],
            raw=c,
            s2126_factors=sorted(
                [(_d(x["from"]), x["factor"])
                 for x in c.get("s2126_consortium_factors", [])]
            ),
        )

    # --- lookups mirroring the workbook semantics ---

    def rpf(self, months_to_off_risk: int) -> float:
        """MATCH(...,1): largest band <= months."""
        f = 0.0
        for m, factor in self.rpf_bands:
            if months_to_off_risk >= m:
                f = factor
        return f

    def leo_debris_rpf(self, months_left: float) -> float:
        """Workbook: IF(AC=0,0, INDEX(S4:S7, MATCH(AC,R4:R7)+1)) — shifted band."""
        if months_left <= 0:
            return 0.0
        idx = 0
        for i, (m, _) in enumerate(self.leo_debris_bands):
            if months_left >= m:
                idx = i
        shifted = idx + 1
        if shifted >= len(self.leo_debris_bands):
            return self.leo_debris_bands[-1][1]  # guard: cap at last band
        return self.leo_debris_bands[shifted][1]

    def equity_pct(self, inception: dt.date, is_consortium: bool) -> float:
        if not is_consortium:
            return 0.0
        return self.equity_by_year.get(inception.year, 0.0)

    def s3123_factor(self, inception: dt.date) -> float:
        """XLOOKUP exact-or-next-smaller on date."""
        f = 0.0
        for d, factor in self.s3123_factors:
            if inception >= d:
                f = factor
        return f

    def s2126_factor(self, inception: dt.date) -> float:
        """S2126 consortium sub-share (XLOOKUP exact-or-next-smaller on date).
        0 before the S2126 participation start (2026-04-01) and when the config
        carries no S2126 schedule."""
        f = 0.0
        for d, factor in self.s2126_factors:
            if inception >= d:
                f = factor
        return f

    def igr_qs_rate(self, entity: str, mapping_code: str, uw_year) -> float:
        for o in self.igr_qs.get("overrides", []) or []:
            if (o["entity"] == entity and o["code"] == mapping_code
                    and int(o["year"]) == int(uw_year or 0)):
                return o["cession"]
        return self.igr_qs["default"].get(entity, 0.0)

    def xol_recovery(self, entity: str, net_of_qs: float) -> float:
        terms = self.igr_xol.get(entity)
        if not terms:
            return 0.0
        if net_of_qs > terms["deductible"]:
            return min(net_of_qs - terms["deductible"], terms["limit"])
        return 0.0
