# Moonlighting Parser

A web-based tool to parse Amion HTML exports and generate trainee moonlighting summaries.

**Live Tool:** [https://sid-dogra.github.io/moonlighting_parser/](https://sid-dogra.github.io/moonlighting_parser/)

## How to Use

### Step 1: Download HTML files from Amion

You need to download two HTML files from Amion:

#### Residency Schedule HTML

![Download Residency Schedule](screenshots/residency_download.png)

#### Contrast Coverage Schedule HTML

![Download Contrast Coverage](screenshots/contrast_download.png)

### Step 2: Upload files to the parser

1. Go to [https://sid-dogra.github.io/moonlighting_parser/](https://sid-dogra.github.io/moonlighting_parser/)
2. Upload your **Trainees.xlsx** file (contains list of residents and fellows)
3. Upload the **Residency Schedule HTML** file
4. Upload the **Contrast Coverage Schedule HTML** file
5. Enter the month name (e.g., "April")
6. Click **Parse and Generate Summary**
7. Download the resulting `.txt` summary file

## Output

The tool generates a summary showing each trainee's shifts and total hours:

```
April Moonlighting Summary
==================================================

John Smith
  04-01-2026 Gramercy weekday           4
  04-05-2026 CBI weekend                8
  Total:                                12

Jane Doe
  04-02-2026 ECEP                       6
  04-08-2026 53rd St weekend            8
  Total:                                14
```

## Privacy

All processing happens locally in your browser. No files are uploaded to any server.
