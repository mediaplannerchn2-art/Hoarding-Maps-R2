# Excel to Exact GPS Site Map - Netlify Edition

A fully static HTML + JavaScript version of the Excel OOH map tool.

## What it does

- Reads `.xlsx`, `.xlsm`, and `.xls` directly in the browser.
- Detects the required columns and common header aliases.
- Validates Latitude and Longitude before generation.
- Uses exact Web-Mercator projection for GPS anchor placement.
- Keeps each arrow tip fixed on the Excel Lat/Long.
- Uses compact collision-aware numbered callouts.
- Draws the full client-facing site table with wrapped Location text.
- Generates 1600x900, 1920x1080, or 2560x1440 PNG output.
- No Python, server, database, login, or Netlify Function is required.

## Required Excel columns

`S.No | City | Media | Area Name | Location | W | H | Lat | Long`

Common alternatives such as `S No`, `Width`, `Height`, `Latitude`, `Longitude`, `Lng`, and `Lon` are supported.

## Netlify deployment - easiest method

1. Extract this ZIP.
2. Go to Netlify and open **Add new site > Deploy manually** / Netlify Drop.
3. Drag the extracted project folder into Netlify.
4. Netlify publishes it immediately. There is no build command.

## Git deployment

- Build command: leave blank
- Publish directory: `.`

`netlify.toml` already contains the publish configuration.

## Important browser/network notes

- Excel parsing uses SheetJS loaded from jsDelivr.
- Map tiles are loaded directly from OpenStreetMap when the user generates a map.
- Internet access is therefore required in the browser.
- The exported image includes OpenStreetMap attribution.
- For heavy commercial usage, use a dedicated tile provider rather than placing high traffic on the public OSM tile service.

## Accuracy

The software does not intentionally move the GPS anchor. The arrow tip is calculated from the Excel Latitude and Longitude using Web Mercator and is projection-checked before drawing. Real-world accuracy still depends on the accuracy of the coordinates supplied in the Excel file and the underlying map data.
