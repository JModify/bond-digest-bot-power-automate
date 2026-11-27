# Digest Bot for Power Automate

Format SharePoint news digests inside Bond Microsoft 365. Power Automate reads the digest, an Excel Office Script applies the Staff News layout, and Bond Outlook sends the formatted copy.

## Set up a flow

| Intake | Where to send the SharePoint digest | Guide |
| --- | --- | --- |
| SDC group mailbox | `sdc@groups.bond.edu.au`. The flow replies to the original Bond sender from `bondstaffnews@bond.edu.au`. | [SDC group mailbox setup](power-automate/group-mailbox.md) |
| Personal mailbox | Your own Bond email address. The flow sends the formatted digest back to your inbox. | [Personal mailbox setup](power-automate/README.md) |

Both guides walk through the Power Automate cards and list the expressions to paste into the designer. Use your Bond account for the connections and workbook.

## Source files

| File | Purpose |
| --- | --- |
| [format-staff-news.ts](power-automate/format-staff-news.ts) | Parses a SharePoint digest and builds the formatted Outlook draft with inline images. |
| [group-post-state.ts](power-automate/group-post-state.ts) | Records the group post IDs that the flow has claimed or processed. |
| [group-flow-definition.json](power-automate/group-flow-definition.json) | Card wiring reference for the group mailbox flow. |
| [flow-definition.json](power-automate/flow-definition.json) | Card wiring reference for the personal mailbox flow. |
| [assets](assets) | Source PNGs embedded in the formatting script. |

Run `npm test` with Node.js 24 or later. The tests use synthetic email and a workbook stub to check formatting, attachments, flow wiring and group post tracking.

Author: Joshua Lavagna-Slater
