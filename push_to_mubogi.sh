#!/usr/bin/env bash
# Push this CV into the Mubogi organisation and create a Release.
#
# WHY YOU MUST RUN THIS YOURSELF:
#   The token available inside the agent workspace is READ-ONLY
#   ("Resource not accessible by integration"). It can clone Mubogi repos but
#   cannot create repos, create branches, or push. A write-capable token is
#   needed, and only you can provide it.
#
# WHAT YOU NEED (pick one):
#   A) A GitHub Personal Access Token (classic) with the "repo" scope, that is
#      a member of the Mubogi organisation with write access. Create at:
#      https://github.com/settings/tokens
#   B) `gh auth login` as an account that can write to the Mubogi org.
#
# USAGE (option A):
#   export GH_WRITE_TOKEN=<your_token>
#   bash push_to_mubogi.sh
#
# USAGE (option B):
#   gh auth login
#   bash push_to_mubogi.sh
#
set -euo pipefail

ORG="Mubogi"
REPO="cv"                 # change if you want a different repo name
BRANCH="main"
TAG="v1.0"

# ---------------------------------------------------------------------------
# Choose the token
# ---------------------------------------------------------------------------
if [[ -n "${GH_WRITE_TOKEN:-}" ]]; then
  export GH_TOKEN="$GH_WRITE_TOKEN"
  PUSH_URL="https://${GH_WRITE_TOKEN}@github.com/${ORG}/${REPO}.git"
  echo "==> Using GH_WRITE_TOKEN."
elif gh auth status >/dev/null 2>&1; then
  echo "==> Using gh CLI credentials."
  PUSH_URL="https://github.com/${ORG}/${REPO}.git"
else
  echo "ERROR: No write credentials. Set GH_WRITE_TOKEN or run 'gh auth login'." >&2
  exit 1
fi

# ---------------------------------------------------------------------------
# Create the repo in the Mubogi org (skips if it already exists)
# ---------------------------------------------------------------------------
echo "==> Creating ${ORG}/${REPO} (skips if it exists)…"
if gh repo view "${ORG}/${REPO}" >/dev/null 2>&1; then
  echo "    Repo already exists."
else
  gh repo create "${ORG}/${REPO}" --public \
    --description "CV - Mubogi Gastavas Jordan" \
    --disable-wiki
fi

# ---------------------------------------------------------------------------
# Push
# ---------------------------------------------------------------------------
echo "==> Pushing ${BRANCH} and tags to ${ORG}/${REPO}…"
git remote remove mubogi >/dev/null 2>&1 || true
git remote add mubogi "${PUSH_URL}"
git push mubogi "${BRANCH}"
git push mubogi "${TAG}"

# ---------------------------------------------------------------------------
# Release with the documents
# ---------------------------------------------------------------------------
echo "==> Creating release ${TAG}…"
gh release create "${TAG}" \
  --repo "${ORG}/${REPO}" \
  "CV_Mubogi_Jordan.pdf" \
  "CV_Mubogi_Jordan.docx" \
  "CV_Master.md" \
  --title "CV v1.0 - Mubogi Gastavas Jordan" \
  --notes "Download my CV:
- CV_Mubogi_Jordan.pdf - designed 3-page CV
- CV_Mubogi_Jordan.docx - editable, ATS-friendly
- CV_Master.md - plain text

Contact: mubogigastavasjordan@gmail.com | +256 754 687 597" \
  || echo "    (Release may already exist - skipping.)"

echo "==> Done."
echo "    Repo:    https://github.com/${ORG}/${REPO}"
echo "    Release: https://github.com/${ORG}/${REPO}/releases/tag/${TAG}"
