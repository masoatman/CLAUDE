"""Buffer ETF mid-period behaviour and portfolio liquidity stress test.

Part 1 prices the option packages behind SMAX (100% buffer) and STEN (10% buffer)
with Black-Scholes plus a simple volatility skew, to estimate what each fund is
worth if you have to SELL it partway through an outcome period during a selloff.

Part 2 stress-tests the illustrative 60/40 from the deck with semi-liquid alts
(or buffer ETFs) funded out of core bonds, and measures liquidity at the trough.

All inputs are assumptions stated below; outputs are illustrative, not forecasts.
"""
from math import exp, log, sqrt
from statistics import NormalDist

N = NormalDist().cdf

# ---- Part 1: option-package pricing ----------------------------------------
R0, Q, ATM0, SKEW = 0.038, 0.012, 0.16, -0.5   # rate, div yield, ATM vol, vol per unit log-moneyness
SMAX_CAP, STEN_CAP = 1.0780, 1.1763            # gross caps, Oct 1 2025 - Sep 30 2026 period
FEE = 0.005


def vol(K, S, atm):
    return max(0.05, atm + SKEW * log(K / S))


def bs(S, K, T, r, q, atm, call):
    if T <= 0:
        return max(S - K, 0) if call else max(K - S, 0)
    v = vol(K, S, atm)
    d1 = (log(S / K) + (r - q + v * v / 2) * T) / (v * sqrt(T))
    d2 = d1 - v * sqrt(T)
    if call:
        return S * exp(-q * T) * N(d1) - K * exp(-r * T) * N(d2)
    return K * exp(-r * T) * N(-d2) - S * exp(-q * T) * N(-d1)


def smax(S, T, r, atm):  # long underlying + long ATM put - short call at cap
    return S * exp(-Q * T) + bs(S, 1, T, r, Q, atm, False) - bs(S, SMAX_CAP, T, r, Q, atm, True)


def sten(S, T, r, atm):  # long underlying + put spread 100/90 - short call at cap
    return (S * exp(-Q * T) + bs(S, 1, T, r, Q, atm, False) - bs(S, 0.90, T, r, Q, atm, False)
            - bs(S, STEN_CAP, T, r, Q, atm, True))


def fund_return(f, S, t_elapsed, r, atm):
    """Fund price return vs inception NAV of 1.

    The flat-rate/simple-skew model misprices the package by ~0.5-1% at inception; that
    gap is removed as an offset that decays to zero at expiry, so held-to-end values
    equal the contractual payoff exactly.
    """
    T = 1 - t_elapsed
    offset = (f(1.0, 1.0, R0, ATM0) - 1) * T
    return (f(S, T, r, atm) - offset) * (1 - FEE * t_elapsed) - 1


def part1():
    print("PART 1 - Value if SOLD mid-period (price return, after pro-rata fee)")
    print(f"{'Scenario':44s}{'S&P':>7s}{'SMAX':>8s}{'STEN':>8s}")
    rows = [
        ("Month 3: -20%, vol 30%, rates -100bp", 0.80, 0.25, R0 - 0.01, 0.30),
        ("Month 6: -20%, vol 30%, rates -100bp", 0.80, 0.50, R0 - 0.01, 0.30),
        ("Month 6: -30%, vol 40%, rates -150bp", 0.70, 0.50, R0 - 0.015, 0.40),
        ("Month 9: -30%, vol 40%, rates -150bp", 0.70, 0.75, R0 - 0.015, 0.40),
        ("Month 6: -20%, vol 28%, rates +150bp (2022)", 0.80, 0.50, R0 + 0.015, 0.28),
        ("Month 6: +10%, vol 14%, rates flat", 1.10, 0.50, R0, 0.14),
        ("Held to end: -30%", 0.70, 1.00, R0, ATM0),
        ("Held to end: -8%", 0.92, 1.00, R0, ATM0),
        ("Held to end: +25%", 1.25, 1.00, R0, ATM0),
    ]
    out = {}
    for name, S, t, r, atm in rows:
        a, b = fund_return(smax, S, t, r, atm), fund_return(sten, S, t, r, atm)
        out[name] = (a, b)
        print(f"{name:44s}{S-1:>7.0%}{a:>8.1%}{b:>8.1%}")
    return out


# ---- Part 2: portfolio liquidity stress --------------------------------------
TARGET_EQ = 0.60
PORTS = {
    "A Base 60/40":            {"eq": .60, "AGG": .25, "TIP": .05, "SGOV": .05, "IAU": .05},
    "B +15% semi-liquid alts": {"eq": .60, "AGG": .10, "TIP": .05, "SGOV": .05, "IAU": .05, "ALT": .15},
    "C +25% semi-liquid alts": {"eq": .60, "AGG": .05, "TIP": .00, "SGOV": .05, "IAU": .05, "ALT": .25},
    "D Buffer sleeve (liquid)": {"eq": .50, "STEN": .10, "AGG": .15, "SMAX": .10, "TIP": .05, "SGOV": .05, "IAU": .05},
}
LIQUID_DEFENSIVE = ("AGG", "TIP", "SGOV", "IAU", "SMAX")   # T+1, typically holds value in a selloff


def part2(p1, withdraw=0.10, fill=0.38):
    """Trough-of-bear liquidity test.

    Sequence at the trough: (1) raise `withdraw` (share of ORIGINAL value) for spending or
    capital calls from liquid defensive assets; (2) rebalance equities back to 60%.
    Semi-liquid alts are assumed marked with a lag (small reported loss) and redeemable only
    via quarterly tenders, pro-rated at `fill` of the amount requested.
    """
    s1_smax, s1_sten = p1["Month 6: -30%, vol 40%, rates -150bp"]
    s2_smax, s2_sten = p1["Month 6: -20%, vol 28%, rates +150bp (2022)"]
    scen = {
        "Deflationary crash (2008/2020-style), trough at month 6":
            {"eq": -.30, "AGG": .04, "TIP": .01, "SGOV": .02, "IAU": .05, "ALT": -.04, "SMAX": s1_smax, "STEN": s1_sten},
        "Inflationary bear (2022-style), trough at month 6":
            {"eq": -.20, "AGG": -.10, "TIP": -.07, "SGOV": .02, "IAU": -.03, "ALT": -.02, "SMAX": s2_smax, "STEN": s2_sten},
    }
    results = {}
    for sname, ret in scen.items():
        print(f"\nPART 2 - {sname}  (withdraw {withdraw:.0%} of original value, then rebalance to 60% equity)")
        print(f"{'Portfolio':26s}{'Port ret':>9s}{'Liquid def':>11s}{'Alt wt':>8s}"
              f"{'Rebal buy':>10s}{'Liquid left':>12s}{'Alt wt after':>13s}{'Alt cash 1Q':>12s}")
        for pname, w in PORTS.items():
            v = {k: x * (1 + ret[k]) for k, x in w.items()}
            tot = sum(v.values())
            liq = sum(v.get(k, 0) for k in LIQUID_DEFENSIVE)
            alt = v.get("ALT", 0)
            eq = v["eq"] + v.get("STEN", 0)
            tot2 = tot - withdraw
            buy = max(0.0, TARGET_EQ * tot2 - eq)
            left = liq - withdraw - buy
            alt_q = alt * fill                      # if the ENTIRE alt position were tendered
            results[(sname, pname)] = dict(ret=tot - 1, liq=liq / tot, alt=alt / tot, buy=buy,
                                           left=left / tot2, alt_after=alt / tot2, alt_q=alt_q)
            print(f"{pname:26s}{tot-1:>9.1%}{liq/tot:>11.1%}{alt/tot:>8.1%}{buy:>10.1%}"
                  f"{left/tot2:>12.1%}{alt/tot2:>13.1%}{(f'{alt_q:.1%}' if alt else '-'):>12s}")
    print("\nColumns: Liquid def = AGG/TIP/SGOV/IAU/SMAX as % of portfolio at trough; Rebal buy = equities"
          " bought to restore 60% (share of original value); Liquid left = liquid defensive remaining"
          " after withdrawal + rebalance; Alt cash 1Q = cash from tendering the whole alt sleeve at a"
          f" {fill:.0%} pro-rata fill.")
    return results


if __name__ == "__main__":
    part2(part1())
