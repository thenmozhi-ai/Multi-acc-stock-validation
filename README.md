# Stock Validation Dashboard (single-file version)

This is the same app as before, but with everything — constants, file detection,
report readers, validation logic, and the Excel exporter — merged into **one
`app.py` file**. There is no `src/` subfolder anymore.

## Why single-file

The multi-file version kept hitting `ModuleNotFoundError: No module named 'src'`
on Streamlit Cloud. That error means the `src/` folder wasn't fully present in
the deployed repo — most commonly because a drag-and-drop GitHub upload silently
skipped it or one of its files. Putting everything in one file removes that
failure mode entirely: there's nothing to lose.

## Deploying

You only need **two files** in your GitHub repo root:

```
your-repo/
├── app.py
└── requirements.txt
```

(`.streamlit/config.toml` is optional — it only sets the app's colour theme.)

1. Go to your GitHub repo.
2. Delete anything left over from the old multi-file attempt (the `src/` folder,
   old `app.py`).
3. Upload this `app.py` and `requirements.txt` to the repo root.
4. On Streamlit Cloud: **Manage app → Reboot app**. If it's a brand-new deploy,
   make sure the **Main file path** is set to `app.py`.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What it does

Same functionality as before — see the in-app help (expand "What files does
this app recognise?" when no files are uploaded) or ask for the feature summary
again if you'd like it repeated here.
