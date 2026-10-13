# XRD 5-Curve Stacked Recipe (verified)

Reference implementation: `E:\codex\测试3\xrd_origin_reference.py`.

## Inputs

- Read-only data, one file per curve, whitespace separated, columns = 2theta, intensity.
- This run used `E:\codex\wang\1-1.txt` .. `5-5.txt`.
- Labels: Ink1..Ink5, black / red / blue / green / purple.

## Steps that matter

- Build one worksheet with 6 columns: x + 5 y series. Write with
  `app.PutWorksheet`, then verify `wks.nrows` and spot-check cells.
- Plot stacked curves: `plotxy iy:=(1,2:6) plot:=200;`.
- Offset each series vertically before writing so the curves do not overlap.
- Hide the legend (`legend.show=0;`); the ink labels act as the legend.
- Per-curve color:
  `layer -p N; set %c -cl <color(r,g,b)>;` - resolve the color number via
  `app.Evaluate("color(r,g,b)")`, never a raw constant.
- (hkl) annotations: `label -a x y -n HK1 "(111)";` - the quotes are required or
  LabTalk swallows the parentheses.
- Reference sticks: `draw -n STK1 -l {x,0,x,h};` with a magenta color.

## Peaks

Use fixed positions, never auto-detection - the low-angle 5-7 deg tail and the
~23/24.6 deg minor peaks get misread as reflections. For this dataset:

- Strong: 25.90, 43.30, 50.44 deg
- Weak: 38.12, 74.20 deg

## Export

```text
expGraph type:=png path:="C:\OriginData" filename:="XRD_5stack_origin" overwrite:=1;
```

Output: `C:\OriginData\XRD_5stack_origin.png`. Do not save `.opju`; the user
presses Ctrl+S.
