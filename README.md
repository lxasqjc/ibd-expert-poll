# IBD Expert Poll

Blinded expert evaluation of AI-generated biological interpretations for Crohn's disease fistula spatial transcriptomics data.

## For Reviewers

Visit the poll at: **https://lxasqjc.github.io/ibd-expert-poll/**

- You can save progress and return later (stored in your browser)
- When finished, click Submit to download your results as JSON
- Email the JSON file to chen.jin1@astrazeneca.net

## For Developers

The poll is a self-contained HTML file (`index.html`). To update responses:

1. Run `assemble_expert_poll.py` (in the main GCCL repo) to populate responses
2. Copy the generated `index.html` here
3. Push to deploy via GitHub Pages

## Study Context

This evaluates whether AI agents augmented with self-learned principles (DiME) produce better biological interpretations than:
- Naive agents (same data, no learned principles)
- Literature-only agents (no data access)

Evaluation is blinded — reviewers do not know which system produced which response.
