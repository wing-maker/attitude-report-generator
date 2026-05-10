# 🎨 TEMPLATE_GUIDE — How to Edit template.pptx

This guide is for designers who want to update the look of generated reports
**without touching any code**.

---

## ✅ What You Can Safely Change

Open `template.pptx` in PowerPoint and freely edit:

- **Colors** — change fills, text colors, accent colors
- **Fonts** — change font family, size, weight
- **Spacing & alignment** — move things around within slides
- **Backgrounds** — add backgrounds, decorative shapes
- **Decorative elements** — add lines, shapes, page numbers, etc.
- **Add/remove slides** at the END (after the Thank You slide) — they'll be ignored by the code

---

## ⚠️ What You Must NOT Change

The code uses two anchors to find things in the template:

### 1. **Placeholder text** like `{CAMPAIGN_NAME}`, `{LIKES}`, `{POST_IMAGE}`

These are literal strings in textboxes. The code does find-and-replace on them.

**Rules:**
- Don't delete or rename them
- Don't add spaces inside the braces (e.g. `{ CAMPAIGN_NAME }` won't work)
- You CAN change the font, color, size of the text containing these placeholders

### 2. **Shape names** with `ph_` prefix (e.g. `ph_client_logo`, `ph_post_image`)

These are anchors used to position images. To see/edit shape names in PowerPoint:
- **Mac**: Home → Arrange → Selection Pane
- **Windows**: Home → Select → Selection Pane

**Rules:**
- Don't delete shapes with `ph_` prefix
- Don't rename them
- You CAN move them, resize them, change fill color

---

## 📋 Full Placeholder Reference

### Cover Slide (slide 1)
| Placeholder | What It Becomes |
|---|---|
| `{CAMPAIGN_NAME}` | Campaign name (e.g. "JOMBELI AMPANG POINT") |
| `{MONTH}` | Campaign month (e.g. "Apr 2026") |
| Shape `ph_client_logo` | Client logo image |
| Shape `ph_hero_image` | Campaign hero image (or removed if user didn't upload) |

### Top 3 Slides (slides 2 & 3)
| Placeholder | What It Becomes |
|---|---|
| `{TOP3_NAME_1}`, `{TOP3_NAME_2}`, `{TOP3_NAME_3}` | Short creator names (inside cards) |
| `{TOP3_FULLNAME_1/2/3}` | Full creator names (below cards) |
| `{TOP3_VIEWS_1/2/3}` | View counts |
| `{TOP3_ENG_1/2/3}` | Engagement counts |
| Shape `ph_top3_image_1/2/3` | Post screenshot images |

### Posting Slides (slides 4-7)
| Placeholder | What It Becomes |
|---|---|
| `{LIKES}`, `{COMMENTS}`, `{VIEWS}`, `{SAVES}`, `{SHARES}` | Engagement numbers |
| `{CREATOR_NAME}` | KOC/KOL name |
| `{POST_URL}` | Link to the post |
| Shape `ph_post_image` | Post screenshot |
| Shape `ph_insight_image` | Insight screenshot |

### Summary Slide (slide 9)
| Placeholder | What It Becomes |
|---|---|
| `{TIKTOK_POSTS}`, `{TIKTOK_LIKES}`, `{TIKTOK_VIEWS}`, etc. | Tiktok totals |
| `{XHS_POSTS}`, `{XHS_LIKES}`, etc. | XHS totals |
| `{IG_POSTS}`, `{IG_LIKES}`, etc. | IG (cross-post) totals |
| `{TOTAL_POSTS}`, `{TOTAL_LIKES}`, etc. | Grand totals |
| `{SUMMARY_PARAGRAPH}` | Auto-generated summary text |

### Thank You Slide (slide 10)
No placeholders — feel free to redesign entirely.

---

## 💡 Tips

- **Test your changes**: After editing, run the tool and generate a report to see how your changes look with real data.
- **Header bar color** is set by the code (from the hero image), not by your template's color.
  If the user doesn't upload a hero image, the default green `#7CB342` is used.
- **Want to change the default green?** Edit `pptx_builder.py`:
  ```python
  DEFAULT_HEADER_COLOR = (124, 179, 66)  # change to your preferred RGB
  ```

---

## 🆘 Common Issues

**"My changes don't appear in generated reports"**
→ Check that you saved `template.pptx` (not "Save As"). The file path matters.

**"Logo / screenshot is in the wrong place"**
→ Move the corresponding `ph_xxx` shape in the template — code uses its position.

**"Text is overlapping placeholders"**
→ Adjust shape positions in PowerPoint. Don't change the placeholder text itself.

**"I want to remove the IG (EXTRA) section"**
→ This is hard-coded in the app logic. Contact the developer instead — but you can edit slide 6
  (the IG divider) to look however you like.

---

## 🎁 Adding More Client Logos

Drop new PNG/JPG files into `brands/` folder:
```
brands/
├── watsons.png       (default)
├── guardian.png      (just add this)
├── caring.png        (and this)
```

The dropdown in the app will automatically include them. No code changes needed.

**Tip**: Use logos that are **transparent PNG** with reasonable proportions (height ~150-300px).
