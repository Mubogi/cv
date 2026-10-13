#!/usr/bin/env bash
# Publish this CV repo to GitHub and create a Release with the documents.
#
# You need a GitHub token that can create repositories (classic PAT with "repo"
# scope, or `gh auth login`). The read-only token inside the agent workspace
# cannot create repos, so run this from YOUR computer or a terminal where you
# are logged in as yourself.
#
# Usage:
#   gh auth login                 # once, if not already logged in
#   bash publish.sh
#
set -euo pipefail

OWNER="MAYI12333"          # change if your username differs
REPO="cv"                  # repo name
BRANCH="main"
TAG="v1.0"

echo "==> Creating repository ${OWNER}/${REPO} (skips if it exists)…"
if gh repo view "${OWNER}/${REPO}" >/dev/null 2>&1; then
  echo "    Repo already exists."
else
  gh repo create "${OWNER}/${REPO}" --public \
    --description "My CV - Mubogi Gastavas Jordan" \
    --source . --remote origin --push
fi

echo "==> Ensuring 'origin' remote points to the repo…"
git remote get-url origin >/dev/null 2>&1 \
  && git remote set-url origin "https://github.com/${OWNER}/${REPO}.git" \
  || git remote add origin "https://github.com/${OWNER}/${REPO}.git"

echo "==> Pushing ${BRANCH} and tags…"
git push -u origin "${BRANCH}"
git push origin "${TAG}"

echo "==> Creating release ${TAG} with the CV documents…"
gh release create "${TAG}" \
  "CV_Mubogi_Jordan.pdf" \
  "CV_Mubogi_Jordan.docx" \
  "CV_Master.md" \
  --title "CV v1.0 — Mubogi Gastavas Jordan" \
  --notes "Download my CV:
- CV_Mubogi_Jordan.pdf — designed 3-page CV
- CV_Mubogi_Jordan.docx — editable, ATS-friendly
- CV_Master.md — plain text

Contact: mubogigastavasjordan@gmail.com | +256 754 687 597"

echo "==> Done."
echo "    Repo:    https://github.com/${OWNER}/${REPO}"
echo "    Release: https://github.com/${OWNER}/${REPO}/releases/tag/${TAG}"
