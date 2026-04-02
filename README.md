# fingerlakes
Here's everything in full:

---

## SKILL.md

**name:** finger-lakes-guide-editor

**description:** Edit, format, and produce the Upstate Finger Lakes Regional Guide as a polished .docx with consistent chapter structure, TLDR sections, color-coded hyperlinks, CTA blocks, and equivalent depth across all five lake chapters. Use this skill whenever Bri asks to work on the Finger Lakes guide, edit a lake chapter, add TLDRs, add hyperlinks, fix consistency, produce the guide doc, or run through the chapter editing workflow. Also trigger for "work on the guide," "next chapter," "edit Cayuga / Seneca / Keuka / Canandaigua / Eastern Lakes," or any request to improve, format, or output the guide. This skill governs ALL Finger Lakes guide editing work — never attempt it without reading this file first.

---

### Purpose

Produce a fully consistent, deeply linked, editorially excellent Finger Lakes Regional Guide as a .docx file. Each chapter must follow the canonical section structure below, carry a TLDR, have color-coded hyperlinks throughout, and close with a CTA block.

**Chapter order:**
1. Cayuga Lake (first chapter to be formatted to the standard — not a pre-existing template)
2. Seneca Lake
3. Keuka Lake
4. Canandaigua Lake
5. The Eastern Lakes
6. Part 1 — Regional Landing Page (last, after chapters are locked)

**CRITICAL ORIENTATION:** The content across all chapters is largely written and strong. This workflow is primarily about ORGANIZATION and CONSISTENCY — applying the canonical section structure, adding TLDRs, adding color-coded hyperlinks, adding CTA blocks, and ensuring equivalent section depth. Do NOT rewrite or substantially edit existing copy. Restructure it, reorder it, and link it. Generate new content only to fill confirmed missing sections (Where to Eat and Where to Stay are explicitly excluded from this pass).

---

### Before Starting Any Chapter

1. Read `references/chapter-template.md`
2. Read `references/hyperlink-rules.md`
3. Read `references/tldr-format.md`
4. Read `references/cta-blocks.md`
5. Read `/mnt/skills/user/upstate-brand-voice/SKILL.md`
6. Read `/mnt/skills/public/docx/SKILL.md`

---

### Workflow Per Chapter

**Step 1 — Audit.** Read existing chapter. Check against canonical structure. Note missing sections, thin sections, inconsistent subheads, missing hyperlinks, missing TLDR, missing CTA.

**Step 2 — Draft.** Write only confirmed missing sections. No em dashes anywhere (overrides brand voice skill default). No exclamation marks. No forbidden phrases.

**Step 3 — Add TLDR.** Place "The long and short of it" immediately before the chapter intro. See tldr-format.md.

**Step 4 — Add Hyperlinks.** Every mention of every linkable entity throughout the chapter. See hyperlink-rules.md.

**Step 5 — Add CTA Block.** Chapter end. See cta-blocks.md.

**Step 6 — Generate .docx.** US Letter (12240 x 15840 DXA), 1-inch margins. Hyperlink colors: Wikipedia `0563C1`, Maps `1A7340`, Upstate `C0392B`, Other `6C3483`. All underlined. Arial throughout.

**Step 7 — Validate and output.** Copy to `/mnt/user-data/outputs/`.

---

### Consistency Rules

| Element | Standard |
|---|---|
| TLDR | Present, before intro, 3-4 sentences |
| At-a-glance table | Length, Depth, Counties, Wine Trail, From NYC, From Buffalo, From Syracuse, Anchor Cities |
| Section order | Per canonical template |
| Heritage section | Present in every chapter with badge track note |
| Wine trail section | Present in all chapters except Eastern Lakes |
| Farms section | Present, separate from Outdoor |
| Outdoor/nature section | Present, separate from Farms |
| Getting here | Drive times table, shore roads, seasonal notes |
| Stamp/badge callout | End of Getting Here |
| CTA block | Chapter end |
| Hyperlinks | Color-coded, every mention |

---

### Content Depth Standards

| Section | Minimum |
|---|---|
| Chapter intro | 3 paragraphs: physical character, Indigenous history, wine/agricultural identity |
| TLDR | 3-4 sentences, action-oriented |
| The Towns | Every named town: population, character note, 1-2 specific details |
| Heritage section | Named sites with full badge list; at least one narrative anchor story |
| Wine trail | How-to-run guidance + Tier 1 producers with full descriptions |
| Farms section | At least 3 named entries |
| Outdoor/nature | At least 3 named sites with named sites list at end |
| Getting here | Drive times table, shore roads, seasonal notes |

---

## references/chapter-template.md

**Chapter structure (in order):**

```
CHAPTER HEADER
Lake Name
Tagline — one sentence, declarative, present tense
Chapter draft · [Counties] · March 2026

IMAGE PLACEHOLDER + caption

TLDR — "The long and short of it"
3-4 sentences. Before the intro.

CHAPTER INTRO — "The lake, and what it keeps producing" [REQUIRED]
P1: Physical character
P2: Indigenous and early settlement history
P3: Wine/agricultural/industrial identity
Optional P4: Connecting thread

AT A GLANCE TABLE [REQUIRED]
Length | Depth | Counties | Wine Trail | From NYC | From Buffalo | From Syracuse | Anchor Cities

THE TOWNS [REQUIRED]
TOWN NAME
COUNTY · POSITION · POP. ~X,XXX
BEST FOR [descriptor]
2-4 paragraphs per town

THE HERITAGE CORRIDOR [REQUIRED]
Open with Indigenous history and its ongoing presence.
Move through reform/settlement/industrial history.
Close with full heritage sites list.
Badge track note: "All sites below are tagged to the Finger Lakes Heritage Trail passport badge."
List format: — [Site Name] [Town] [Brief description.]

THE [LAKE] WINE TRAIL [CONDITIONAL — all lakes except Eastern]
Open paragraph: terroir/thermal character
How to run the trail paragraph
Producers: Tier 1 (full descriptions) / Tier 2 (2 sentences) / Tier 3 (list, 1 sentence)

Tier 1 format:
[Producer Name]
Subtitle: [5-7 word evocative line]
Short description: [2 sentences]
Description: [3-5 sentences — history, winemaker, what to drink, logistics]

FARMS, CIDER, AND FOOD [REQUIRED — SEPARATE FROM OUTDOOR]
Named farms, cideries, farmers markets, food producers.
2-4 sentences each. Badge track noted.

OUTDOOR AND NATURE [REQUIRED — SEPARATE FROM FARMS]
Do not combine with Farms section.
Primary draw: 1 paragraph.
Secondary sites: 2-4 sentences each.
Named sites list at end.

GETTING HERE AND GETTING AROUND [REQUIRED]
Opening: "[Lake name] requires a car."
Airports paragraph.
Drive times table: Origin | Destination | Time
Shore roads paragraph.
Seasonal notes paragraph.

STAMP/BADGE CALLOUT [REQUIRED]
Wine trail stamps: X producers tagged to [Trail Name] badge. Count toward Finger Lakes Wine master badge.
Heritage stamps: [sites] — all Finger Lakes Heritage Trail badge.
Nature stamps: [sites] — Finger Lakes Outdoor badge.
[Farm Trail / Craft Beverage Trail if applicable]

CTA BLOCK [REQUIRED]
See cta-blocks.md.
```

**Section flexibility notes:**
- Heritage Corridor and Haudenosaunee History can be split for Cayuga and Canandaigua (richest content)
- Eastern Lakes has no wine trail section — use "What these lakes share" connective content instead
- At a Glance table: place after intro when physical stats carry narrative weight (Seneca); place before intro for shorter lakes

---

## references/hyperlink-rules.md

**Color system:**

| Type | Color | Hex | When to use |
|---|---|---|---|
| Wikipedia | Blue | `0563C1` | Named historical figures, places/events/institutions with Wikipedia articles |
| Google Maps | Green | `1A7340` | Any named navigable location: parks, towns, restaurants, wineries, farms, museums, landmarks |
| Upstate | Red | `C0392B` | Other lake chapters, main FL page, named itineraries, trail collection pages |
| Other | Purple | `6C3483` | Official websites, NPS pages, state park pages, winery/trail official sites |

**Scope rule:** Link every mention, not just first mention.

**Dual-link rule:** Named wineries, restaurants, farms, museums, and parks get two links per mention: purple on the name (official site) + green on the town/location immediately following (Google Maps).

**Upstate URLs:**
- Cayuga chapter: `https://upstate.tourismo.app/itineraries/cayuga-lake`
- Seneca chapter: `https://upstate.tourismo.app/itineraries/seneca-lake`
- Keuka chapter: `https://upstate.tourismo.app/itineraries/keuka-lake`
- Canandaigua chapter: `https://upstate.tourismo.app/itineraries/canandaigua-lake`
- Eastern Lakes chapter: `https://upstate.tourismo.app/itineraries/eastern-lakes`
- Main FL page: `https://upstatebound.com/guides/finger-lakes-region-04b4ec42-84de-4c38-aa0a-f689dc88d7a6`
- The Freedom Line: `https://upstate.tourismo.app/itineraries/the-freedom-line`
- Around Seneca: `https://upstate.tourismo.app/itineraries/around-seneca`
- Gorge Country: `https://upstate.tourismo.app/itineraries/gorge-country`
- Cayuga Farm Loop: `https://upstate.tourismo.app/itineraries/cayuga-farm-loop`
- Cayuga Wine Trail: `https://upstate.tourismo.app/trails/cayuga-lake-wine-trail`
- Seneca Wine Trail: `https://upstate.tourismo.app/trails/seneca-lake-wine-trail`
- Keuka Wine Trail: `https://upstate.tourismo.app/trails/keuka-lake-wine-trail`
- Canandaigua Wine Trail: `https://upstate.tourismo.app/trails/canandaigua-lake-wine-trail`
- Heritage Trail: `https://upstate.tourismo.app/trails/finger-lakes-heritage-trail`
- Farm Trail: `https://upstate.tourismo.app/trails/finger-lakes-farm-trail`
- Craft Beverage Trail: `https://upstate.tourismo.app/trails/finger-lakes-craft-beverage-trail`

**Google Maps format:** `https://www.google.com/maps/search/?api=1&query=[Place+Name+URL+encoded]`

**Key Wikipedia URLs:**
- Harriet Tubman: `.../wiki/Harriet_Tubman`
- Elizabeth Cady Stanton: `.../wiki/Elizabeth_Cady_Stanton`
- Frederick Douglass: `.../wiki/Frederick_Douglass`
- William H. Seward: `.../wiki/William_H._Seward`
- Konstantin Frank: `.../wiki/Konstantin_Frank`
- Hermann Wiemer: `.../wiki/Hermann_J._Wiemer_Vineyard`
- Sullivan-Clinton Campaign: `.../wiki/Sullivan%E2%80%93Clinton_campaign`
- Haudenosaunee Confederacy: `.../wiki/Iroquois`
- Cayuga Nation: `.../wiki/Cayuga_Nation`
- Seneca Nation: `.../wiki/Seneca_Nation_of_Indians`
- Treaty of Canandaigua: `.../wiki/Treaty_of_Canandaigua`
- Elizabeth Blackwell: `.../wiki/Elizabeth_Blackwell`
- Glenn Curtiss: `.../wiki/Glenn_Curtiss`
- Millard Fillmore: `.../wiki/Millard_Fillmore`
- Susan B. Anthony: `.../wiki/Susan_B._Anthony`
- Red Jacket: `.../wiki/Red_Jacket_(Seneca_leader)`
- Public Universal Friend: `.../wiki/Public_Universal_Friend`
- Finger Lakes AVA: `.../wiki/Finger_Lakes_AVA`
- Watkins Glen SP: `.../wiki/Watkins_Glen_State_Park`
- Taughannock Falls: `.../wiki/Taughannock_Falls_State_Park`
- Women's Rights NHP: `.../wiki/Women%27s_Rights_National_Historical_Park`

**TBD placeholder:** Use `https://TBD` for unconfirmed URLs. Known TBDs: Finger Lakes Heritage Trail official site, Finger Lakes Farm Trail official site, Finger Lakes Craft Beverage Trail official site.

---

## references/tldr-format.md

**Placement:** Immediately before "The lake, and what it keeps producing." After header and image placeholder.

**Format:** Bold label "The long and short of it" at Heading3. Followed directly by the TLDR paragraph.

**Length:** 3-4 sentences. No more.

**Tone:** Action-oriented. The lake's identity + the primary reason to visit + practical orientation. Written as the most interesting person at the table leaning over and saying what you actually need to know.

**Avoid:** No em dashes. No exclamation marks. No forbidden phrases. Don't restate the tagline. Don't summarize the whole chapter — pick the two or three things that matter most.

**Examples:**

Cayuga: *Cayuga is the longest lake in the region and the one that rewards people who stay more than two nights. The reform corridor between Auburn and Seneca Falls is the most historically consequential fifteen miles in New York State. Base in Ithaca or Aurora, run the east shore on day two, and plan at least one morning at the Farmers Market before you do anything else.*

Seneca: *Seneca is the deepest lake entirely within New York State, and its thermal mass is the reason the wine trail here has more producers than any other loop in the region. Run the west shore on day one and the east shore on day two: Route 414 is quieter, steeper, and produces wines that earn the Mosel comparison they're always being given. Get to Watkins Glen before nine in the morning or after four in the afternoon.*

Keuka: *Keuka is the only Y-shaped Finger Lake, and the shape gives it something no other trail has: vineyard visible on three sides at once from the right elevation. American wine started here. So did American flight. One day covers the complete circuit; Hammondsport is the right anchor for both stories.*

Canandaigua: *Canandaigua is the most approachable lake in the region: smaller wine trail, less traffic, a downtown that functions without wine-trail branding pressing down on it. The Treaty of Canandaigua was signed at the north end in 1794 and is still honored annually on November 11. Come in late September for the Naples Grape Festival, or any other time when you want the Finger Lakes to feel like a place people actually live.*

Eastern Lakes: *The eastern Finger Lakes are the ones nobody packaged, and that is the whole argument for them. Skaneateles has the clearest water in the region and a village that has been doing exactly what it does for two centuries. Owasco gives you Auburn and everything that comes with it. Hemlock and Canadice are the wildest land in the Finger Lakes, which happened entirely by accident.*

---

## references/cta-blocks.md

**Placement:** Very end of each chapter, after the stamp/badge callout. Last content element before next chapter.

**Format:** Light shading (fill `EAF4FB`), red border all sides, bold header "Explore more of the Finger Lakes."

**Structure:** (1) Cross-chapter navigation sentence with red Upstate links. (2) Main guide link sentence. (3) Optional related itinerary links.

**Chapter-specific copy:**

**Cayuga:** Cayuga is the longest lake and the most historically layered. When you're ready to go deeper into the region, the [Seneca Lake chapter] picks up the wine story thirty-eight miles west, and the [Finger Lakes regional guide] has the full picture. [Plan your visit on Upstate →] / Related: [The Freedom Line Heritage Itinerary] · [The Cayuga Farm Loop]

**Seneca:** Seneca is the wine center. The [Keuka Lake chapter] is where the wine story actually started, twenty miles west on a Y-shaped lake above Hammondsport. The [Cayuga Lake chapter] picks up the reform corridor. The [Finger Lakes regional guide] connects all five. [Plan your visit on Upstate →] / Related: [Around Seneca Wine Itinerary]

**Keuka:** Keuka is the origin. The [Seneca Lake chapter] picks up the wine story as it matured: thirty-three producers on the deepest lake in the region. The [Finger Lakes regional guide] has the full picture across all five lakes. [Plan your visit on Upstate →]

**Canandaigua:** Canandaigua is the westernmost and the most approachable. The [Seneca Lake chapter] is twenty miles east when you're ready for a bigger wine trail. The [Finger Lakes regional guide] connects the full region. [Plan your visit on Upstate →]

**Eastern Lakes:** The eastern lakes are the Finger Lakes before the brand arrived. When you're ready for the wine country that built the region's reputation, the [Cayuga Lake chapter] and the [Seneca Lake chapter] are the next stops. The [Finger Lakes regional guide] has the full picture. [Plan your visit on Upstate →]

---

## references/chapter-audit.md

**Cayuga Lake — First chapter to format.** Content strong. Reorganize into canonical structure, add TLDR, hyperlinks, CTA. Tasks: TLDR, hyperlinks, CTA, split farms/outdoor if combined, confirm at-a-glance has all 8 fields, reformat stamp callout, standardize image placeholders.

**Seneca Lake.** Strong chapter. Tasks: TLDR, hyperlinks, CTA, split farms/outdoor, confirm producer count (33 vs 28 discrepancy — fix throughout), confirm at-a-glance fields, reformat stamp callout.

**Keuka Lake.** Good structure. Tasks: TLDR, hyperlinks, CTA, split farms/outdoor, reconcile wine trail count (shows 6, lists 9+), reformat stamp callout.

**Canandaigua Lake.** Good historical content. Tasks: TLDR, hyperlinks, CTA, split farms/outdoor, confirm at-a-glance fields, reformat stamp callout.

**Eastern Lakes.** Good character writing. Tasks: TLDR, hyperlinks, CTA, organize Beak & Skiff etc. into a Farms section, consolidate outdoor content into a named Outdoor section, reformat stamp callout (Heritage and Outdoor only, no wine trail), keep existing "What these lakes share" connective content in place of wine trail section.

**Part 1 — Regional Landing Page.** Edit last, after chapters locked. Consistency and hyperlinking only. Tasks: TLDR, hyperlinks on all lake names (red), itineraries (red), historical figures (blue), named locations (green), named trails (red to trail collection pages), verify Wine Enthusiast 2025 claim before publishing.
