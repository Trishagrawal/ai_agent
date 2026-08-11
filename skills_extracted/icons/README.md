# Google Icons — AFG deck icon set (compliance replacement)

Replaces the previous AI-generated icon stock. Every glyph below is sourced
directly from **Google Fonts Material Symbols** (https://fonts.google.com/icons),
licensed under the Apache License 2.0 — free to use, modify, and redistribute,
including commercially, with no attribution required (though noted here for audit purposes).

Source repo: https://github.com/google/material-design-icons
Style used: Material Symbols Outlined, regular weight.

## Format — matches the skill's existing icons/ spec exactly
- 400×400 px, transparent background PNG
- `icons/light/<name>.png` — navy (#001438) glyph, for light-themed slides
- `icons/dark/<name>.png` — white (#ffffff) glyph, for dark-themed slides
- Same filenames as the original set, so this folder is a drop-in replacement
  (just overwrite the existing `icons/light/` and `icons/dark/` folders in the skill package)

## Note on styling
The original AI-generated set used a two-tone navy+cyan accent per icon. Google's
Material Symbols are single-color/monochrome glyphs, so these are single-tone
(navy on light, white on dark) rather than two-tone. This keeps every icon
guaranteed-compliant (unmodified Google glyph, recolored only) rather than
partially hand-edited. If you want the cyan accent look back on specific icons,
that would need manual per-icon touch-up on top of these.

## Concept → Google icon mapping

| Filename | Google Material Symbol used |
|---|---|
| `automotive-battery.png` | `battery_charging_full` |
| `automotive-car.png` | `directions_car` |
| `automotive-ev-charging.png` | `ev_station` |
| `automotive-steering-wheel.png` | `car_repair` |
| `automotive-warning.png` | `warning` |
| `checklist.png` | `checklist` |
| `corporate-briefcase.png` | `work` |
| `corporate-clock.png` | `schedule` |
| `corporate-lamp.png` | `table_lamp` |
| `corporate-monitor.png` | `desktop_windows` |
| `corporate-phone.png` | `call` |
| `cost.png` | `savings` |
| `data.png` | `database` |
| `finance.png` | `account_balance` |
| `financial-services-atm.png` | `local_atm` |
| `financial-services-credit-card.png` | `credit_card` |
| `financial-services-pie-chart.png` | `pie_chart` |
| `financial-services-trend-chart.png` | `monitoring` |
| `financial-services-wallet.png` | `account_balance_wallet` |
| `governance.png` | `gavel` |
| `growth.png` | `trending_up` |
| `mobile.png` | `smartphone` |
| `operations.png` | `settings` |
| `productivity.png` | `bolt` |
| `real-estate-bed.png` | `bed` |
| `real-estate-building.png` | `apartment` |
| `real-estate-for-rent.png` | `real_estate_agent` |
| `real-estate-house.png` | `house` |
| `real-estate-key.png` | `key` |
| `retail.png` | `shopping_bag` |
| `retail-cart.png` | `shopping_cart` |
| `retail-delivery-truck.png` | `local_shipping` |
| `retail-gift.png` | `redeem` |
| `retail-storefront.png` | `storefront` |
| `retail-tag.png` | `sell` |
| `roadmap.png` | `timeline` |
| `security.png` | `security` |
| `speed.png` | `speed` |
| `teamwork.png` | `groups` |
| `technology.png` | `memory` |

You can preview/search any of these names directly at https://fonts.google.com/icons