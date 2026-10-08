// The script behind the Claude Code and Codex tiles on the Connect your AI page: curl -fsSL <sheets>/connect/claude | sh
// Adapted from agent-kanban's connect script: it prints each command before it runs it, and is safe to run again.
export const APPS = ['claude', 'codex'];
export const safeHost = (host) => (/^https?:\/\/[a-z0-9.-]+(:\d{1,5})?$/i.test(String(host ?? '')) ? String(host) : null);
// ---------- the script ----------

// A POSIX sh script with these sheets' address baked in. It finds Claude Code and Codex, adds Sheets,
// and starts sign-in. It prints each command before running it, and is safe to run again.
export function connectScript({ host, app = null, name = 'Team' }) {
  const mcp = `${host}/mcp`;
  const slug = (new URL(host).hostname + new URL(host).pathname).replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toLowerCase().slice(0, 40);
  const apps = app ? [app] : APPS;
  const label = String(name).replace(/[^\w .&+-]/g, '').slice(0, 60) || 'Team';
  return `#!/bin/sh
# Connects ${apps.map((a) => (a === 'claude' ? 'Claude Code' : 'Codex')).join(' and ')} to ${label} on wOS Sheets (${host}).
# It prints each command before it runs it, and it is safe to run again.
#   Read it first:  curl -fsSL ${host}/connect${app ? `/${app}` : ''}
#   Dry run:        curl -fsSL ${host}/connect${app ? `/${app}` : ''} | sh -s -- --dry-run
set -eu

URL='${mcp}'
NAME='wos-sheets'
ALT='wos-sheets-${slug}'
APPS='${apps.join(' ')}'
DRY=0
DONE_IN=''

say() { printf '%s\\n' "$*"; }
show() { printf '  $ %s\\n' "$*"; }
run() { show "$@"; if [ "$DRY" = 0 ]; then "$@"; fi; }
# Sign-in opens a browser and may need the keyboard: give it the terminal, not this script's input.
run_tty() { show "$@"; if [ "$DRY" = 0 ]; then if (: </dev/tty) 2>/dev/null; then "$@" </dev/tty; else "$@"; fi; fi; }
added() { DONE_IN="\${DONE_IN:+$DONE_IN and }$1"; }

# Which name to use: ours if it is free or already points here, otherwise one with this server's address in it.
pick() {
  for n in "$NAME" "$ALT"; do
    out=$("$@" "$n" 2>/dev/null) || { CHOSEN=$n; EXISTS=0; return; }
    case "$out" in *"$URL"*) CHOSEN=$n; EXISTS=1; return ;; esac
  done
  say "  Both $NAME and $ALT are taken by other servers. Remove one and run this again."
  exit 1
}

claude_code() {
  say ''
  say 'Claude Code'
  pick claude mcp get
  if [ "$EXISTS" = 1 ]; then say "  Already added as $CHOSEN."
  else run claude mcp add --transport http --scope user "$CHOSEN" "$URL"; fi
  if [ "$DRY" = 1 ] && [ "$EXISTS" = 0 ]; then
    say '  If it asks you to sign in, a browser window opens to sign in. Sign in there.'
    show claude mcp login "$CHOSEN"
  elif claude mcp get "$CHOSEN" 2>/dev/null | grep -q 'Needs authentication'; then
    say '  Signing in: a browser window opens to sign in. Sign in there.'
    run_tty claude mcp login "$CHOSEN"
  else say '  Connected.'; fi
  added 'Claude Code'
}

codex_cli() {
  say ''
  say 'Codex'
  pick codex mcp get
  if [ "$EXISTS" = 1 ]; then say "  Already added as $CHOSEN."
  else
    say '  If it asks you to sign in, a browser window opens to sign in. Sign in there.'
    run_tty codex mcp add "$CHOSEN" --url "$URL"
  fi
  if [ "$DRY" = 0 ] && codex mcp list 2>/dev/null | grep "^$CHOSEN " | grep -q 'Not logged in'; then
    say '  Signing in: a browser window opens to sign in. Sign in there.'
    run_tty codex mcp login "$CHOSEN"
  elif [ "$EXISTS" = 1 ]; then say '  Connected.'; fi
  added 'Codex'
}

main() {
  for arg in "$@"; do
    case "$arg" in
      -n|--dry-run) DRY=1 ;;
      *) say "Unknown option: $arg (the only option is --dry-run)"; exit 2 ;;
    esac
  done
  say 'Connecting ${label} (wOS Sheets) at '"$URL"
  [ "$DRY" = 1 ] && say 'Dry run: this shows what would run and changes nothing.'
  for app in $APPS; do
    case "$app" in
      claude) if command -v claude >/dev/null 2>&1; then claude_code; fi ;;
      codex) if command -v codex >/dev/null 2>&1; then codex_cli; fi ;;
    esac
  done
  if [ -z "$DONE_IN" ]; then
    say ''
    case "$APPS" in
      claude) say 'Claude Code is not installed here. Get it at https://claude.com/claude-code, then run this again.' ;;
      codex) say 'Codex is not installed here. Get it at https://developers.openai.com/codex, then run this again.' ;;
      *) say 'Neither Claude Code nor Codex is installed here. Install one, then run this again.' ;;
    esac
    exit 1
  fi
  say ''
  if [ "$DRY" = 1 ]; then say 'Dry run finished. Run it without --dry-run to connect.'
  else say "Done. Open $DONE_IN (restart it if it is open) and ask it to open your sheets"; fi
}

main "$@"
`;
}

