# Excel Model Exporter  
**Using LLMs to Review and Understand Excel Models**

---

## Why This Exists

One of the most effective uses of LLMs is helping people get up to speed on unfamiliar code.

Excel models don’t share that benefit:

- logic is scattered across sheets,
- formulas are hidden behind cell references,
- charts obscure dependencies,
- and sharing `.xlsm` files is often risky or impractical.

This project addresses that gap.

---

## The Idea

Instead of asking an LLM to build or modify a spreadsheet, this tool focuses on something simpler and more reliable:

> **Export an Excel model into a clean, text-based representation that humans and LLMs can review.**

The result is a deterministic, read-only artifact that captures structure, logic, and visualization bindings — without opening the original workbook.

---

## What It Exports

Running the exporter produces a folder containing:

- Sheet layouts (structure)
- Formulas (logic)
- Charts (metadata, not images)
- Named ranges
- VBA modules (if permitted)
- A brief model summary and usage guide

All outputs are plain text and safe to share.

---

## Example

The `/examples` folder contains a public export from a simple OHLC data visualization model.

The model is intentionally basic. The goal is not sophistication — it’s to demonstrate how quickly a reviewer (human or LLM) can understand how a spreadsheet works.

---

## What This Is Not

This tool is **not**:

- a trading system
- a strategy generator
- a backtester
- an Excel replacement

It makes no predictions and executes nothing.  
Its sole purpose is **review and understanding**.

---

## Typical Workflow

1. Build or modify an Excel model  
2. Run the exporter  
3. Review the exported artifact (yourself, with collaborators, or with an LLM)  
4. Return to Excel only after structure and logic are clear  

---

## VBA Security Note

Exporting VBA requires enabling:


If disabled, VBA extraction is skipped automatically.

---

## Inspiration

This project was inspired by reading about how LLMs have been used to help new hires get up to speed on large codebases — not by replacing engineers, but by accelerating understanding.

The same idea applies to Excel.
