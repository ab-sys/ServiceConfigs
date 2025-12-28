# ~/.bashrc
# Executed by bash(1) for non-login shells.

# --- Guard: only for interactive shells --------------------------------------
case $- in
  *i*) ;;
  *) return ;;
esac

# --- History & shell behavior ------------------------------------------------
HISTCONTROL=ignoreboth
HISTSIZE=1000
HISTFILESIZE=2000
shopt -s histappend
shopt -s checkwinsize

# --- Debian chroot indicator (optional) -------------------------------------
if [ -z "${debian_chroot:-}" ] && [ -r /etc/debian_chroot ]; then
  debian_chroot="$(cat /etc/debian_chroot)"
fi

# --- Colors (prompt & ls) ----------------------------------------------------
# Use tput if available, fall back to ANSI.
__bashrc_has_tput() { command -v tput >/dev/null 2>&1; }
__bashrc_supports_color() {
  __bashrc_has_tput && tput setaf 1 >/dev/null 2>&1
}

# Color escape sequences for PS1 need \[ \] so Bash can calculate line length.
RED='\[\033[1;31m\]'
GREEN='\[\033[1;32m\]'
YELLOW='\[\033[1;33m\]'
BLUE='\[\033[1;34m\]'
CYAN='\[\033[1;36m\]'
RESET='\[\033[0m\]'

# --- Prompt ------------------------------------------------------------------
# Keep the prompt logic in one place. Avoid multiple PS1 assignments.
__bashrc_set_prompt() {
  local host='\H' user='\u' cwd='\w' time='\A'
  local chroot=''
  if [ -n "${debian_chroot:-}" ]; then
    chroot="(${debian_chroot}) "
  fi

  # Prefer color prompt when supported.
  if __bashrc_supports_color; then
    PS1="${YELLOW}${time} ${GREEN}${user}@${BLUE}${host}:${CYAN}${cwd} ${RESET}\$ "
  else
    PS1="${chroot}${user}@${host}:${cwd}\$ "
    return
  fi

  # Prepend chroot if present (kept uncolored to avoid nesting complexity).
  if [ -n "$chroot" ]; then
    PS1="${chroot}${PS1}"
  fi

  # Set xterm title for common terminals.
  case "${TERM:-}" in
    xterm*|rxvt*|screen*|tmux*)
      PS1="\[\e]0;${chroot}${user}@${host}: ${cwd}\a\]${PS1}"
      ;;
  esac
}
__bashrc_set_prompt

# --- Core aliases ------------------------------------------------------------
if command -v dircolors >/dev/null 2>&1; then
  if [ -r "$HOME/.dircolors" ]; then
    eval "$(dircolors -b "$HOME/.dircolors")"
  else
    eval "$(dircolors -b)"
  fi
  alias ls='ls --color=auto'
  alias grep='grep --color=auto'
  alias fgrep='fgrep --color=auto'
  alias egrep='egrep --color=auto'
fi

alias ll='ls -lh --color=auto'
alias la='ls -lah --color=auto'

# --- Docker helpers ----------------------------------------------------------
# Only define docker aliases if docker exists.
if command -v docker >/dev/null 2>&1; then
  alias dps='docker ps --format "table {{.Names}}\t{{.Image}}\t{{.Status}}\t{{.Ports}}"'
  alias dlog='docker logs -f'
  alias dexec='docker exec -it'
  alias dimages='docker images --format "table {{.Repository}}\t{{.Tag}}\t{{.Size}}"'
fi

# --- Readline: history search with up/down ----------------------------------
bind '"\e[A": history-search-backward'
bind '"\e[B": history-search-forward'

# --- Editor ------------------------------------------------------------------
export EDITOR="${EDITOR:-vim}"

# --- User extensions ---------------------------------------------------------
# Keep host-specific or personal additions in ~/.bash_aliases (not in this file).
if [ -r "$HOME/.bash_aliases" ]; then
  . "$HOME/.bash_aliases"
fi

# --- Bash completion ---------------------------------------------------------
if ! shopt -oq posix; then
  if [ -r /usr/share/bash-completion/bash_completion ]; then
    . /usr/share/bash-completion/bash_completion
  elif [ -r /etc/bash_completion ]; then
    . /etc/bash_completion
  fi
fi
