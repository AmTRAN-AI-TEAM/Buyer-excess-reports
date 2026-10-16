from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class CustomerProfile:
    code: str
    separate_price_file: bool = False
    price_headers: tuple[str, ...] = ("Uprice",)
    preferred_cost_sheet: str | None = None


CUSTOMER_PROFILES = {
    "AVTC": CustomerProfile(code="AVTC", preferred_cost_sheet="成本价"),
    "RAKEN": CustomerProfile(
        code="RAKEN",
        separate_price_file=True,
        price_headers=("USD",),
    ),
}


def profile_for(customer: str | None) -> CustomerProfile | None:
    if customer is None:
        return None
    return CUSTOMER_PROFILES.get(customer.upper())
