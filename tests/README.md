# Local redesign regression checks

Run from the repository root. These checks never commit, push, or deploy.

```sh
node tests/logic.test.js
python3 tests/check-migration.py
python3 tests/run-browser-tests.py
python3 tests/check-docx.py /tmp/habeas-v3-implementation-tests
python3 tests/run-ui-tests.py
```

The browser runners require existing Google Chrome and copies of the exact pinned browser scripts. No package installation is used. Defaults use these files under `/tmp`:

- `habeas-review-pizzip-3.1.7.js`
- `habeas-review-docxtemplater-3.50.0.js`
- `habeas-review-file-saver-2.0.5.js`
- `habeas-review-jszip-3.10.1.js`
- `habeas-review-docx-preview-0.3.3.js`

The scripts are copies of the CDN URLs already listed in `index.html`; they are test inputs rather than dependency changes. Use `--pizzip`, `--vendor-dir`, `--chrome`, `--output`, and `--artifacts` to choose other local paths. If Chrome or loopback networking is blocked by a sandbox, run the browser runner with the environment's approval mechanism.

`logic.test.js` checks language behavior, not JavaScript execution of templates. `browser.test.js` compiles real DOCX packages, inspects Word XML, saves diagnostic documents, exercises status ownership with deliberately reordered completions, and verifies Save/Load and Studio inspection. The browser runner also parses every XML/rels part and checks ZIP integrity.

`check-migration.py` compares the sample against the recorded starting commit, verifies unchanged Word structures/non-document parts, and restricts prose differences to approved syntax and defined-term substitutions.

`check-docx.py` verifies content-type coverage, relationship targets and IDs, style/numbering references, cell contents, footnote reference IDs, and footnote own markers. It is a targeted package checker, not a complete OOXML schema validator.

`ui.test.js` runs the real application and Studio in iframes using the pinned preview library. Its runner binds an ephemeral server **only to 127.0.0.1**, replaces CDN URLs in test responses with local vendor copies, and shuts down the server afterward. Downloads are captured as Blobs for inspection. This is local testing, not publishing.

Results and representative DOCX files go to `/tmp/habeas-v3-implementation-tests` by default. Microsoft Word on Windows inspection is a separate, required acceptance step. Automated success does not establish Word acceptance.
