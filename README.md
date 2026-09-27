# Weather Windows for Marine Operations

This repository contains the source for the three-slide master's-level teaching deck:

`Weather_Windows_Marine_Operations.pptx`

## Generate the PowerPoint

```bash
python -m pip install -r requirements.txt
python generate_weather_windows_pptx.py
```

The script creates the requested file in the repository directory. The slides explain weather windows, present operation-specific weather criteria using the DNV N001 basis, and show a simple offshore lift case with a visual time line.

> Note: DNV N001 requires operation-specific criteria; the example numerical limits in the deck are illustrative teaching values and must not replace the approved project design basis or marine procedure.
