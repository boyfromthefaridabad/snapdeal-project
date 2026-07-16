# E-Commerce Pricing, Risk & Trust Analytics — Power BI Project

**Dataset:** `Ecommerce_Dataset.xlsx` (sheets: `Products`, `DailyScrapeLog`)
**Model:** Star schema — `Products` (dimension + attributes) ↔ `DailyScrapeLog` (fact, one row per Product per ScrapingDate) joined on `ProductID`.

Create these two base relationships in Power BI:
- `Products[ProductID]` (1) → `DailyScrapeLog[ProductID]` (*)
- Mark `DailyScrapeLog[ScrapingDate]` as a Date table, or build a separate `Calendar` table and relate it to `DailyScrapeLog[ScrapingDate]`.

---

## Task 1 — Pricing vs Satisfaction Exception Detection

**Goal:** Flag products that are simultaneously overpriced, under-rated, high-risk, and under-selling relative to their category.

### Measures

```dax
Avg Price in Category =
CALCULATE (
    AVERAGE ( Products[MRP] ),
    ALLEXCEPT ( Products, Products[Category] )
)

Avg Rating in Category =
CALCULATE (
    AVERAGE ( Products[Rating] ),
    ALLEXCEPT ( Products, Products[Category] )
)

Median Sales in Category =
CALCULATE (
    MEDIANX ( ALLEXCEPT ( Products, Products[Category] ), Products[Sales_Last30Days] )
)

Risk Score =
Products[AvgDiscount_Last30Days]
    * ( 1 - ( Products[Rating] / 5 ) )
    * Products[StockValue]

Risk Score Percentile Rank =
RANKX (
    ALL ( Products ),
    [Risk Score],
    ,
    DESC
) / COUNTROWS ( ALL ( Products ) )

Is Exception Product =
VAR PriceHigh   = Products[MRP] > [Avg Price in Category]
VAR RatingLow   = Products[Rating] < [Avg Rating in Category]
VAR RiskTop20   = [Risk Score Percentile Rank] <= 0.20
VAR SalesBelowMedian = Products[Sales_Last30Days] < [Median Sales in Category]
RETURN
IF ( PriceHigh && RatingLow && RiskTop20 && SalesBelowMedian, 1, 0 )
```

Build the table visual using: `ProductName`, `Category`, `MRP`, `[Avg Price in Category]`,
`Rating`, `[Avg Rating in Category]`, `Risk Score`, `Sales_Last30Days`,
`[Median Sales in Category]`, filtered where `[Is Exception Product] = 1`.
*(All logic lives in measures — no calculated columns needed, satisfying the "measures only" constraint.)*

### Business interpretation

These products fail on **every dimension simultaneously**: customers are paying a premium
for something rated worse than category peers, the capital tied up in that inventory
is disproportionately risky, and sales aren't compensating for any of it. This is not
one bad metric — it is a *compounding* failure pattern, which is why it needs a table,
not a single KPI.

**Recommended manager actions, by root cause:**
- **Re-price** — if the rating is only slightly below average and reviews mention "value for money," the price is the real problem; a targeted markdown, not a blanket discount, restores competitiveness without triggering another quality-perception hit.
- **Quality improvement** — if the rating gap is large and return-rate is elevated, the product itself is broken; further discounting will only accelerate loss-making, low-satisfaction sales.
- **Removal / delisting** — if both price and rating gaps are severe *and* stock value is high, the product is actively destroying margin and brand trust at scale — delist and liquidate remaining stock rather than continue merchandising it.

---

## Task 2 — Promotion Effectiveness Trend Analysis

**Goal:** Separate "real" discounting during promotions from a product's normal, everyday discount level, using the daily scrape history.

### Measures

```dax
Baseline Discount (30d non-promo) =
VAR NonPromoDates =
    CALCULATETABLE (
        VALUES ( DailyScrapeLog[ScrapingDate] ),
        DailyScrapeLog[IsPromotional] = 0
    )
VAR Last30NonPromo =
    TOPN ( 30, NonPromoDates, DailyScrapeLog[ScrapingDate], DESC )
RETURN
CALCULATE (
    AVERAGE ( DailyScrapeLog[DiscountPercentage] ),
    KEEPFILTERS ( Last30NonPromo ),
    DailyScrapeLog[IsPromotional] = 0
)

Daily Discount =
AVERAGE ( DailyScrapeLog[DiscountPercentage] )

Promotional Lift =
VAR CurrentDiscount = [Daily Discount]
VAR Baseline        = [Baseline Discount (30d non-promo)]
RETURN
IF ( ISBLANK ( CurrentDiscount ) || ISBLANK ( Baseline ), BLANK(), CurrentDiscount - Baseline )

Is Promo Day (Any) =
IF ( CALCULATE ( MAX ( DailyScrapeLog[IsPromotional] ) ) = 1, 1, 0 )
```

### Visual
Line chart, X-axis = `ScrapingDate`, lines = `[Daily Discount]`, `[Baseline Discount (30d non-promo)]`,
and `[Promotional Lift]` (or plot `Promotional Lift` as a column overlay to show spikes clearly).
Add a category/product slicer since baseline discount is product-specific.

### Business interpretation

The **baseline** answers "what does this product normally cost after typical discounting?"
The **lift** isolates the *incremental* discount attributable to the promotion itself.

- If promotional lift is consistently large and coincides with a genuine sales-volume spike (cross-check against `UnitsSold`) → the promotion is generating **real incremental demand**.
- If the "baseline" price was already quietly creeping up in the days just before a promotion (a classic **anchoring/inflate-then-discount pattern**), the "big discount" is partly an illusion — the customer isn't actually paying less than they would have a few weeks earlier.
- This distinction matters because promotions that only *simulate* savings erode customer trust once discovered, and inflate reported promotional ROI without real margin or demand benefit.

---

## Task 3 — Discount vs Rating Causation Analysis

**Goal:** Test whether discounting *causes* better ratings, or whether the apparent relationship is confounded by sales/review volume.

### Measures

```dax
Avg Discount =
AVERAGE ( DailyScrapeLog[DiscountPercentage] )

Avg Rating =
AVERAGE ( Products[Rating] )

Review Count =
SUM ( Products[ReviewCount] )

Rating Controlled for Review Volume =
VAR ReviewBucket =
    SWITCH (
        TRUE(),
        Products[ReviewCount] < 50, "Low (<50)",
        Products[ReviewCount] < 500, "Medium (50-500)",
        "High (500+)"
    )
RETURN
CALCULATE (
    AVERAGE ( Products[Rating] ),
    ALLEXCEPT ( Products, Products[ReviewCount] ),   -- placeholder; bucket via calc column or grouping table recommended
    FILTER ( ALL ( Products ), TRUE() )
)

-- Cleaner approach: create a "Review Volume Band" grouping (Modeling > New Group, or a
-- small DAX table) with bands Low / Medium / High, then:

Avg Rating by Review Band =
CALCULATE ( AVERAGE ( Products[Rating] ) )   -- sliced by the Review Volume Band on the axis

Correlation Discount-Rating (approx) =
VAR N = COUNTROWS ( Products )
VAR SumX = SUMX ( Products, Products[AvgDiscount_Last30Days] )
VAR SumY = SUMX ( Products, Products[Rating] )
VAR SumXY = SUMX ( Products, Products[AvgDiscount_Last30Days] * Products[Rating] )
VAR SumX2 = SUMX ( Products, Products[AvgDiscount_Last30Days] ^ 2 )
VAR SumY2 = SUMX ( Products, Products[Rating] ^ 2 )
VAR Numerator = N * SumXY - SumX * SumY
VAR Denominator = SQRT ( ( N * SumX2 - SumX ^ 2 ) * ( N * SumY2 - SumY ^ 2 ) )
RETURN DIVIDE ( Numerator, Denominator )
```

### Visuals
1. **Scatter plot**: X = `[Avg Discount]`, Y = `Products[Rating]`, one dot per product, with Power BI's
   built-in trend line turned on (Format pane → Analytics → Trend line) for the regression line.
2. **Supporting scatter**: X = `[Avg Discount]`, Y = `[Review Count]` — shows whether discounting drives review *volume* rather than review *quality*.
3. **Bar/line chart**: Rating by **Review Volume Band** (Low/Medium/High) — shows whether rating differences persist once review count is controlled for.

### Business interpretation

- If discount and rating show a mild positive slope **only** in the raw scatter, but that
  slope **flattens or disappears** once you compare within the same review-volume band,
  the apparent "discounts improve satisfaction" relationship is a **confound**: heavier
  discounting drives more purchases → more reviews → and high-volume products regress
  toward an average (often higher-looking) rating simply through statistical averaging,
  not genuinely better products.
- If discount correlates strongly with **review count** but only weakly with **rating itself**,
  that's direct evidence ratings are improving mainly through volume dynamics (more buyers,
  more "average" reviews diluting outliers) rather than the product actually satisfying
  customers more.
- Conclusion for the manager: don't treat "discounted products have higher ratings" as
  proof that discounting improves quality perception — it's more likely a sampling effect.

---

## Task 4 — Financially Risky Inventory KPI

### Measures

```dax
Risk Score =
Products[AvgDiscount_Last30Days]
    * ( 1 - ( Products[Rating] / 5 ) )
    * Products[StockValue]

Total Stock Value =
SUM ( Products[StockValue] )

At-Risk Stock Value =
CALCULATE (
    SUM ( Products[StockValue] ),
    FILTER ( Products, [Risk Score] > 0 )   -- or use a materiality threshold, see note
)

-- Recommended: define "at risk" as Risk Score above a materiality threshold
-- (e.g., top quartile of Risk Score) rather than simply >0, since >0 is almost every row.

Risk Threshold (P75) =
PERCENTILEX.INC ( ALL ( Products ), [Risk Score], 0.75 )

At-Risk Stock Value (P75+) =
SUMX (
    FILTER ( Products, [Risk Score] >= [Risk Threshold (P75)] ),
    Products[StockValue]
)

% Inventory Value At Risk =
DIVIDE ( [At-Risk Stock Value (P75+)], [Total Stock Value] )
```

### KPI Card with conditional formatting
Use the `% Inventory Value At Risk` measure on a **Card** visual, then add conditional
formatting via a background-color rule (or use the "KPI" visual / a gauge):
- **Green**: `< 10%`
- **Yellow**: `10% – 25%`
- **Red**: `> 25%`

(Implement with a measure returning a color hex/state, bound to the card's conditional formatting, e.g. a `Risk KPI Color` measure using `SWITCH(TRUE(), [%...] < 0.10, "#2ECC71", [%...] <= 0.25, "#F1C40F", "#E74C3C")`.)

### Business interpretation

The Risk Score deliberately multiplies three failure modes together: **how heavily a
product is being discounted**, **how poorly it's rated**, and **how much capital is tied
up in it as inventory**. A product can be heavily discounted *and* fine if it's well
rated and low-stock; the KPI only flags capital that is both being given away in
discounts *and* attached to a badly-reviewed item. A rising red KPI signals the company
is simultaneously (a) mispricing — because heavy discounting was needed to move stock —
and (b) carrying quality problems at scale, both of which directly erode gross margin
and long-run brand equity if unaddressed.

---

## Task 5 — Trust-Weighted Rating Index

### Measures

```dax
Trust Score =
Products[Rating]
    * LOG ( 1 + Products[ReviewCount] )
    * ( 1 - Products[ReturnRate] )

Trust-Weighted Avg Rating =
DIVIDE (
    SUMX ( Products, Products[Rating] * [Trust Score] ),
    SUMX ( Products, [Trust Score] )
)

Simple Avg Rating =
AVERAGE ( Products[Rating] )

Review-Weighted Avg Rating =
DIVIDE (
    SUMX ( Products, Products[Rating] * Products[ReviewCount] ),
    SUM ( Products[ReviewCount] )
)
```

### Visual
A clustered bar/column chart with three columns — `Simple Avg Rating`,
`Review-Weighted Avg Rating`, `Trust-Weighted Avg Rating` — as a small comparison
panel (or three KPI cards side by side), optionally sliced by Category.

### Business interpretation

A product can rack up a high **simple average rating** from a handful of reviews with no
returns data behind it — that number is fragile and easy to game. `LOG(1 + ReviewCount)`
rewards ratings that are backed by real sample size but with **diminishing returns**
(so a product with 10,000 reviews doesn't dominate the index purely on volume the way a
plain review-weighted average would). Multiplying by `(1 − ReturnRate)` penalizes products
customers rated well *before* they actually used them, then sent back — a common pattern for
items that look good on the page but disappoint in person. The result: a product with a
perfect 5.0 star rating but only 4 reviews and a 30% return rate is *less trustworthy* than
a 4.3-star product with 3,000 reviews and a 4% return rate — and the Trust-Weighted Index
surfaces that, while a naive star-rating leaderboard would not.

---

## Task 6 — Context-Aware Price Banding

### DAX (Calculated Column *is required here* since the task explicitly asks for one; percentiles still recompute dynamically at the visual/filter level via the paired measures below)

```dax
-- Calculated column (static classification for reference / conditional formatting):
Price Band =
VAR P30 =
    PERCENTILEX.INC ( ALLSELECTED ( Products ), Products[MRP], 0.30 )
VAR P70 =
    PERCENTILEX.INC ( ALLSELECTED ( Products ), Products[MRP], 0.70 )
RETURN
SWITCH (
    TRUE(),
    Products[MRP] <= P30, "Low",
    Products[MRP] <= P70, "Medium",
    "High"
)
```

> Note: a true calculated **column** cannot react to slicers at query time (columns are
> computed once, at refresh). To make the band genuinely filter-context-aware as the task
> requires, pair it with measures that recompute the thresholds live and re-bucket via
> measure logic in the visual:

```dax
Dynamic P30 Threshold =
PERCENTILEX.INC ( ALLSELECTED ( Products ), Products[MRP], 0.30 )

Dynamic P70 Threshold =
PERCENTILEX.INC ( ALLSELECTED ( Products ), Products[MRP], 0.70 )

Dynamic Price Band (Measure) =
VAR AvgPriceSelected = AVERAGE ( Products[MRP] )
RETURN
SWITCH (
    TRUE(),
    AvgPriceSelected <= [Dynamic P30 Threshold], "Low",
    AvgPriceSelected <= [Dynamic P70 Threshold], "Medium",
    "High"
)

Product Count by Dynamic Band =
VAR P30 = [Dynamic P30 Threshold]
VAR P70 = [Dynamic P70 Threshold]
RETURN
SWITCH (
    TRUE(),
    SELECTEDVALUE ( 'Price Band Table'[Band] ) = "Low",
    COUNTROWS ( FILTER ( ALLSELECTED ( Products ), Products[MRP] <= P30 ) ),
    SELECTEDVALUE ( 'Price Band Table'[Band] ) = "Medium",
    COUNTROWS ( FILTER ( ALLSELECTED ( Products ), Products[MRP] > P30 && Products[MRP] <= P70 ) ),
    COUNTROWS ( FILTER ( ALLSELECTED ( Products ), Products[MRP] > P70 ) )
)
```

Create a small disconnected `'Price Band Table'` with one column `Band` = {"Low","Medium","High"} to drive the axis, and use `Product Count by Dynamic Band` as the value — this is what makes the bucket counts recalculate live from whatever category/brand/time slicer is applied.

### Visual
Stacked column or donut chart of product count (or stock value) by `Dynamic Price Band`,
with Category and Brand slicers on the page so the recalculation is visibly demonstrated.

### Business interpretation

A **fixed** price band (e.g., "Low = under ₹500") makes no sense across categories that
span wildly different price scales — ₹500 is expensive for Grocery but trivially cheap for
Electronics, so a static band mislabels almost everything outside the category it was
tuned for. Percentile-based bands solve this by asking "cheap/expensive **relative to
what's currently in view**" — so filtering to "Electronics only" recomputes the 30th/70th
percentile *within* Electronics prices, correctly reclassifying products that looked
"High" against the whole catalog but are actually mid-range within their own category.
This filter-context sensitivity is exactly what makes DAX percentile measures more honest
than hardcoded thresholds for real merchandising decisions — a manager filtering into a
single category needs bands calibrated to that category, not to the whole catalog.

---

## Suggested Report Pages
1. **Overview** — KPI cards (Total Products, Total Stock Value, % Inventory At Risk)
2. **Exception Products** (Task 1) — table + slicers
3. **Promotion Trends** (Task 2) — line chart with product/category slicer
4. **Discount vs Rating** (Task 3) — scatter + supporting scatter + banded bar chart
5. **Risk KPI** (Task 4) — KPI card, conditional formatting
6. **Trust Index** (Task 5) — comparison bar chart
7. **Price Banding** (Task 6) — dynamic band chart with slicers
