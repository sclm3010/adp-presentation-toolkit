# Value-Iteration ADP Presentation Toolkit

This repository contains the source scripts and final presentation artifacts for
a technical report on value-iteration adaptive dynamic programming (ADP), based
on Wei, Liu, and Lin, *IEEE Transactions on Cybernetics* (2016).

## Deliverables

- `outputs/ADP_论文汇报_10页思路版.pptx`: ten-slide Chinese presentation.
- `outputs/ADP论文汇报讲稿.docx`: matching presentation script.
- `build_adp_outline_10slides.mjs`: final outline-oriented deck generator.
- `build_adp_speech_docx.py`: Word script generator.
- Earlier deck iterations are retained to show the design process.

## Topics covered

- Bellman/value-iteration updates for nonlinear discrete-time systems.
- Convergence and admissibility conditions.
- Critic/action neural-network implementation.
- Interpretation of the paper's simulation results.

## Regenerating the artifacts

The Word script requires Python 3.10+ and `python-docx`:

```powershell
python -m pip install -r requirements.txt
python build_adp_speech_docx.py
```

The slide generators were built for the Codex desktop artifact runtime. Install
the JavaScript dependency first:

```powershell
pnpm install
node build_adp_outline_10slides.mjs
```

They locate the bundled runtime under `~/.cache/codex-runtimes` by default. Set
`CODEX_RUNTIME_ROOT` to the runtime's `dependencies` directory when it is stored
elsewhere.

## Scope

This is a paper-explanation and presentation-engineering project. It does not
claim to be a full numerical reproduction of the paper's control experiments.
