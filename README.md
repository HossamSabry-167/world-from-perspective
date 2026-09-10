# World From Perspective — Word-powered blog

A small, free, self-hosted blog. Write your posts in Word (Arabic works
properly there, unlike Blogger's editor), upload the `.docx` file
through GitHub's website, and the site rebuilds itself automatically —
no coding, no server, no monthly cost.

**How it works:** you drop a `.docx` file into the `content/` folder on
GitHub. A robot (GitHub Actions) notices, converts it to a web page
with the images placed where they were in the document, and publishes
it. Arabic posts are automatically shown right-to-left with a
readable Arabic font; anything else is shown left-to-right.

---

## One-time setup (15 minutes)

### 1. Create a free GitHub account
Go to [github.com/signup](https://github.com/signup) if you don't
already have an account.

### 2. Create a new repository
- Click the **+** in the top right → **New repository**.
- Name it anything, e.g. `world-from-perspective`.
- Set it to **Public** (required for free GitHub Pages).
- Click **Create repository**. Leave it empty for now.

### 3. Upload these project files
- On your new repository's page, click **Add file → Upload files**.
- Drag in *everything* from this project folder, keeping the folder
  structure (`content/`, `docs/`, `assets/`, `templates/`, `scripts/`,
  `.github/`, `config.json`, `requirements.txt`, `README.md`).
  GitHub's uploader preserves folder paths when you drag a whole
  folder in from your file manager, or you can upload the ZIP and use
  GitHub's "unzip" — if that option isn't available in your browser,
  upload each folder one at a time by dragging it in.
- Commit directly to the `main` branch.

### 4. Let it build once
- Go to the **Actions** tab of your repository. You should see a
  "Build blog" run start automatically (or click **Run workflow** if
  it didn't). Wait for the green checkmark (~30 seconds).
- If it fails with a permissions error, go to **Settings → Actions →
  General → Workflow permissions**, choose **Read and write
  permissions**, save, then re-run the workflow from the Actions tab.

### 5. Turn on GitHub Pages
- Go to **Settings → Pages**.
- Under "Build and deployment", set **Source** to **Deploy from a
  branch**, **Branch** to `main`, folder to **/docs**. Save.
- After a minute, GitHub shows your live URL, something like:
  `https://your-username.github.io/world-from-perspective/`

That's it — your blog is live (with a "no posts yet" message).

---

## Publishing a post

1. Write your article in Word, images and all, exactly as you want it
   to look. Use Word's **Heading 1** style for the title if you want
   it pulled in automatically (or just set the document's Title under
   File → Info → Properties).
2. On GitHub, open the `content/` folder → **Add file → Upload
   files** → drop your `.docx` → commit.
3. Wait about 30–60 seconds for the Actions tab to finish (a small
   yellow dot turns into a green check).
4. Refresh your blog — the new post is there, newest posts first.

**Updating a post:** upload a file with the exact same name again — it
replaces the old version.
**Removing a post:** delete the `.docx` from `content/` on GitHub and
commit.
**Custom order/date:** rename the file with a date prefix, e.g.
`2026-03-04-my-trip.docx`.

## Customizing

Edit `config.json` directly on GitHub (click the file → pencil icon)
to change:
- `site_title`, `site_tagline`, `footer_text`
- `accent_color` (a hex color used for links and highlights)

Saving the file triggers a rebuild automatically.

## Notes and limits

- Images are placed in the same order they appear in your Word
  document (this is how the conversion tool reads the file); very
  specific pixel layouts, text-wrapped images, or multi-column Word
  layouts won't be reproduced exactly, but photos, headings, bold/
  italic, lists, and quotes all come through cleanly.
- Arabic vs. non-Arabic is detected automatically per post from the
  text itself — no setting needed.
- This is entirely free: GitHub Pages hosting and GitHub Actions
  build minutes have no cost for a public repository at blog scale.
- If you'd rather use a custom domain name later (e.g.
  `www.yourblog.com` instead of the github.io address), GitHub Pages
  supports that for free too — you'd just need to own the domain and
  point its DNS at GitHub; ask if you want help with that step later.
