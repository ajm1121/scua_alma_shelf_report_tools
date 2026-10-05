"""Settings that are meant to be edited between shelf-reading projects."""

# How far a shelf position can drift from call-number order before the
# student workbook recommends a move. The written procedure uses 20.
# The staff-report template formula uses 10. Pass --threshold to override.
ORDER_THRESHOLD = 20

# Extra places students mark off while searching for missing items.
# The item's Alma location is always listed on its own.
SEARCH_LOCATIONS = [
    "Rare Double Oversize",
    "Rare Double",
    "Rare Oversize",
    "Rare LC",
    "Oregon Collection",
    "Reference",
]
