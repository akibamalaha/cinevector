# CineVector — Movie Intelligence Engine

Netflix-styled Streamlit dashboard for the SE488 RAG/FAISS project.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

The first run will download `sentence-transformers/all-MiniLM-L6-v2` and
`google/flan-t5-base` from Hugging Face — this needs internet access and
takes a minute or two. After that they're cached.

## Deploy on Streamlit Community Cloud

Same flow as your `fpro-demo` deployment:

1. Push this folder (`app.py`, `requirements.txt`,
   `movie_plots_premium_filled_partial.xlsx`) to a GitHub repo.
2. On https://share.streamlit.io, point a new app at `app.py` in that repo.
3. Deploy. First boot will be slow (model downloads) — subsequent loads
   are fast since Streamlit caches the models/indexes for the session.

## What's inside

- **Hero banner** — a featured movie with a generated cinematic backdrop, synopsis,
  and a "Shuffle Featured" button, similar to a real streaming-service hero.
- **Real interactive carousel** — not a decorative CSS marquee: hover-reveal left/right
  arrows that actually scroll, and a genuine Autoplay toggle (setInterval-based, loops
  at the end), same behavior as a native Netflix row.
- **Movie detail modal** — click "More Info" (in the hero, under the carousel, or under
  any poster in Explore) to open a real `st.dialog` popup with poster, characters,
  themes, keywords, and full synopsis.
- **Explore** — dataset preview, genre/year/plot-length distributions, poster grid
  with per-movie "Info" buttons
- **Vector Observatory** — PCA projection of the 384-D embeddings
- **Index Lab** — all 5 FAISS index types (Flat, IVF, PQ, IVFPQ, HNSW)
  with a live "Retrieval Arena" comparing recall@5 (vs exact Flat) and
  latency on any query you type
- **Data Studio** — SQL queries (SQLite, in-memory), semantic search, and
  a hybrid SQL-filter + semantic-rank mode
- **Evidence Engine** — the RAG chat, using the same hybrid retriever
  (exact-entity match + semantic + TF-IDF + structured + concept scoring)
  and FLAN-T5 prompt logic from your notebook, with retrieved evidence
  shown under every answer

- **Manage Data** — upload your own `.xlsx`/`.csv` (replace or append), add/edit/delete
  rows directly in an inline spreadsheet editor, and export the current dataset back
  out as CSV or Excel. Any change here rebuilds embeddings, all 5 FAISS indexes, and
  TF-IDF automatically the next time another tab needs them.

All retrieval/generation logic is ported directly from
`SE488_RAG_Project_FINAL_UPDATED.ipynb` — same scoring weights, same
FAISS configs (`nlist=5`, PQ `m=8,bits=4`, HNSW `M=32`), same
question-type-aware prompts and grounded fallback.
