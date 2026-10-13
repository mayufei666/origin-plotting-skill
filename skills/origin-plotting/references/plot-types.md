# Plot Types (verified on this machine)

| Goal | Plot ID / call | Verified |
| --- | --- | --- |
| Line | 201 | yes |
| Line + Symbol | 202 | yes |
| Scatter | 200 | yes |
| Column / Bar | 203 | yes |
| Histogram | 303 | not tested |
| Stacked Column | 310 | NO - produced an empty graph here, do not use |

Call form:

```text
plotxy iy:=(xCol,yCol1:yColN) plot:=<ID>;
```

## Notes

- `203` is the correct ID for bar and column charts on this Origin build. `300` and
  `310` silently produced empty plots, so only use what the table marks verified.
- Set column types before plotting: x column `wks.colN.type=1;`, y columns
  `wks.colN.type=2;`.
- Set axis titles with `xb.text$=` and `yl.text$=`; add units in parentheses.
- Show the legend (`legend.show=1;`) only when there are multiple series and it will
  not overlap other labels; otherwise hide it.

## Setup tips

- Label axes with units, for example `Time (s)`, `Intensity (a.u.)`.
- Use a log axis for data spanning several decades: `layer.x.log=1;`.
- Export at 300 dpi or higher for publication.
