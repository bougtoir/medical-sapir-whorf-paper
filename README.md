# The Nosological Imperative: Disease Naming as Therapeutic Intervention in Pain Medicine

**疾患命名の治療的意義：疼痛医学における医療版サピア＝ウォーフ仮説**

## Overview

This project contains a perspective paper proposing the **Medical Sapir-Whorf Hypothesis**: that nosological categories (disease names and classifications) do not merely describe clinical reality but actively constitute it, with genuinely therapeutic effects in pain medicine.

### Key Arguments

1. **Medical Sapir-Whorf Hypothesis**: Disease naming shapes clinical cognition, patient experience, and healthcare system behavior through five channels: perceptual channeling, diagnostic foreclosure, patient ontological validation, research infrastructure alignment, and policy/resource allocation.

2. **Pain Medicine as Ideal Test Case**: Chronic pain was historically under-classified in ICD-10. The ICD-11 chronic pain chapter (MG30), nociplastic pain, and chronic primary pain represent natural experiments in therapeutic naming.

3. **Nosological Therapeutic Loop (NTL)**: A formal model describing the self-reinforcing cycle through which disease naming generates therapeutic effects: Naming → Validation → Attention → Research → Treatment → Refinement.

4. **Boundary Conditions**: Distinguishes therapeutic naming from disease mongering along three dimensions: origin, evidence base, and outcome trajectory.

## Output Files

| File | Description |
|------|-------------|
| `output/manuscript_en.docx` | English manuscript (perspective article) |
| `output/manuscript_ja.docx` | Japanese manuscript (展望論文) |
| `output/figures_tables_en.pptx` | Editable figures and tables (English) |
| `output/figures_tables_ja.pptx` | Editable figures and tables (Japanese) |

## Generation Scripts

```bash
# Generate English manuscript
python scripts/create_manuscript_en.py

# Generate Japanese manuscript
python scripts/create_manuscript_ja.py

# Generate figures/tables PPTX (both EN and JA)
python scripts/create_figures_pptx.py
```

## Dependencies

```bash
pip install python-docx python-pptx
```

## Structure

- `scripts/create_manuscript_en.py` — English docx generation
- `scripts/create_manuscript_ja.py` — Japanese docx generation
- `scripts/create_figures_pptx.py` — Editable PPTX figures and tables (EN + JA)
- `output/` — Generated output files

## References

28 references in Vancouver style, covering:
- Linguistic relativity (Sapir, Whorf, Winawer et al.)
- Philosophy of classification (Hacking, Kleinman)
- ICD-11 chronic pain classification (Treede, Barke, Nugraha et al.)
- Nociplastic pain (Kosek, Nijs, Fitzcharles et al.)
- Diagnostic labeling effects (O'Keeffe, Zadro et al.)
- Disease mongering (Moynihan et al.)
