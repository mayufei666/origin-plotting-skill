---
name: origin-plotting
description: Drive the user's already-open OriginLab Origin (OriginPro) over COM to build, style, and export figures on this Windows machine - XRD stacked patterns, bar/column, line, scatter, and multi-series plots. Use when the user asks to plot or chart data with Origin, or to reproduce or edit an Origin figure here.
metadata:
  short-description: Plot data in the local Origin via COM
---

# Origin Plotting (this machine, verified)

Origin is installed and normally already open. Drive the running instance; do not
launch a second copy and do not kill the existing process.

## Connect (verified here)

- Call `pythoncom.CoInitialize()` first.
- Reuse the open instance with
  `win32com.client.GetActiveObject("Origin.ApplicationSI")`, and only fall back to
  `win32com.client.Dispatch("Origin.ApplicationSI")`.
- `Origin.ApplicationSI` is the working ProgID on this machine.
- Do NOT use the `originpro` package here; it hangs.
- Do NOT trust `Origin.Application`; it is unreliable here.
- Set `app.Visible = True` so the user can see and edit the figure.

## Hard rules

- Treat source data as read-only. Never rewrite the user's data files.
- Do NOT save `.opju` from the script; COM save can crash the connection. Tell the
  user to press Ctrl+S in Origin.
- Export with LabTalk `expGraph`, for example
  `expGraph type:=png path:="C:\OriginData" filename:="fig" overwrite:=1;`
  Do NOT use the clipboard-based MCP `export_graph`; it fails here.
- After exporting, check the PNG exists and is larger than 1000 bytes, then show it
  to the user.

## Reliable mechanics

| Task | Working approach |
| --- | --- |
| Write data | `app.PutWorksheet(page, rows, 0, 0)` with `rows` a list of row lists, then read back `wks.nrows` and spot-check cells |
| Column labels | `col(i)[L]$="name";` |
| Curve color | `layer -p N; set %c -cl <n>;` where `<n>` comes from `app.Evaluate("color(r,g,b)")` |
| Legend | `legend.show=0;` to hide, `legend.show=1;` to show |
| Text annotation | `label -a x y -n NAME "text";` - quote the text or LabTalk swallows parentheses |
| Axis titles | `xb.text$="X"; yl.text$="Y";` |
| Axis range | `layer.x.from=..; layer.x.to=..;` |
| Reference sticks | `draw -n NAME -l {x,y0,x,y1};` |

Never assign a color by writing a raw number straight to a plot property; it fails
silently. Always resolve it through `color(r,g,b)`.

## Workflow

1. Confirm Origin is open, then connect as above.
2. Inspect the data first: path, delimiter, columns, units, and ranges. Then pick a
   plot type - see [references/plot-types.md](references/plot-types.md).
3. Build the worksheet, write the data, and read it back to verify.
4. Plot, then style axes and annotations. Show a legend only when it will not collide
   with other labels.
5. Export the PNG with `expGraph`, verify the file, and display it.
6. Report the graph name and PNG path, and remind the user to press Ctrl+S.

## Ready-made driver

For a generic chart, use [scripts/origin_plot.py](scripts/origin_plot.py):

```text
python scripts/origin_plot.py --data FILE --x-col 1 --y-cols 2,3 --plot-type bar \
  --out C:\OriginData\fig.png --xlabel "Category" --ylabel "Value"
```

Supported `--plot-type`: `line`, `scatter`, `line+symbol`, `bar`.

For the verified XRD stacked-pattern recipe, see
[references/xrd-recipe.md](references/xrd-recipe.md).
