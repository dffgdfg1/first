#!/usr/bin/env bash
set -euo pipefail

TARGET_LOCALE="zh_CN.UTF-8"
SET_TIMEZONE=0
INSTALL_INPUT_METHOD=1
ASSUME_YES=0
ORIGINAL_ARGS=("$@")

usage() {
  cat <<'USAGE'
Ubuntu Simplified Chinese setup script

Usage:
  bash ubuntu-zh-cn-setup.sh [options]

Options:
  -y, --yes            Auto-confirm apt installation
      --no-ime         Do not install the Fcitx5 Chinese input method
      --timezone       Also set the system timezone to Asia/Shanghai
  -h, --help           Show help

Installs:
  - Simplified Chinese language packs: language-pack-zh-hans, language-pack-gnome-zh-hans
  - Locale: zh_CN.UTF-8
  - Chinese fonts: Noto CJK, WenQuanYi
  - Input method: Fcitx5 + Chinese addons

Log out or reboot after installation so the desktop environment can reload language settings.
USAGE
}

log() {
  printf '\033[1;32m[zh-cn]\033[0m %s\n' "$*"
}

warn() {
  printf '\033[1;33m[warn]\033[0m %s\n' "$*" >&2
}

die() {
  printf '\033[1;31m[error]\033[0m %s\n' "$*" >&2
  exit 1
}

while (($#)); do
  case "$1" in
    -y|--yes)
      ASSUME_YES=1
      ;;
    --no-ime)
      INSTALL_INPUT_METHOD=0
      ;;
    --timezone)
      SET_TIMEZONE=1
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      die "Unknown option: $1"
      ;;
  esac
  shift
done

if [[ ${EUID} -ne 0 ]]; then
  log "Root permission is required. Re-running with sudo..."
  exec sudo -E bash "$0" "${ORIGINAL_ARGS[@]}"
fi

if [[ ! -r /etc/os-release ]]; then
  die "Cannot read /etc/os-release, so the system type cannot be detected."
fi

# shellcheck disable=SC1091
. /etc/os-release

if [[ "${ID:-}" != "ubuntu" ]]; then
  warn "Detected ${PRETTY_NAME:-unknown}. This script is designed for Ubuntu, but will continue."
fi

if ! command -v apt-get >/dev/null 2>&1; then
  die "apt-get was not found. This script only supports Ubuntu/Debian-like systems."
fi

APT_FLAGS=()
if [[ ${ASSUME_YES} -eq 1 ]]; then
  APT_FLAGS=(-y)
fi

REQUIRED_PACKAGES=(
  language-pack-zh-hans
  locales
  fonts-noto-cjk
)

OPTIONAL_BASE_PACKAGES=(
  language-pack-gnome-zh-hans
  language-pack-kde-zh-hans
  fonts-noto-cjk-extra
  fonts-wqy-microhei
  fonts-wqy-zenhei
)

IME_PACKAGES=(
  fcitx5
  fcitx5-chinese-addons
  fcitx5-config-qt
  fcitx5-frontend-gtk2
  fcitx5-frontend-gtk3
  fcitx5-frontend-gtk4
  fcitx5-frontend-qt5
  im-config
)

log "Updating apt package indexes..."
apt-get update

package_exists() {
  apt-cache show "$1" >/dev/null 2>&1
}

collect_available_packages() {
  local package
  for package in "$@"; do
    if package_exists "${package}"; then
      printf '%s\n' "${package}"
    else
      warn "Package ${package} was not found in apt sources; skipped."
    fi
  done
}

AVAILABLE_REQUIRED_PACKAGES=()
for package in "${REQUIRED_PACKAGES[@]}"; do
  if package_exists "${package}"; then
    AVAILABLE_REQUIRED_PACKAGES+=("${package}")
  else
    die "Required package ${package} was not found. Check apt sources and retry."
  fi
done

mapfile -t AVAILABLE_OPTIONAL_BASE_PACKAGES < <(collect_available_packages "${OPTIONAL_BASE_PACKAGES[@]}")

log "Installing Simplified Chinese language packs and fonts..."
apt-get install "${APT_FLAGS[@]}" "${AVAILABLE_REQUIRED_PACKAGES[@]}" "${AVAILABLE_OPTIONAL_BASE_PACKAGES[@]}"

if [[ ${INSTALL_INPUT_METHOD} -eq 1 ]]; then
  mapfile -t AVAILABLE_IME_PACKAGES < <(collect_available_packages "${IME_PACKAGES[@]}")
  if [[ ${#AVAILABLE_IME_PACKAGES[@]} -gt 0 ]]; then
    log "Installing the Fcitx5 Chinese input method..."
    apt-get install "${APT_FLAGS[@]}" "${AVAILABLE_IME_PACKAGES[@]}"
  else
    warn "No Fcitx5 packages were found in apt sources; input method installation was skipped."
  fi
fi

log "Generating ${TARGET_LOCALE} locale..."
if grep -qE "^[#[:space:]]*${TARGET_LOCALE//./\\.}[[:space:]]+UTF-8" /etc/locale.gen; then
  sed -i -E "s|^[#[:space:]]*(${TARGET_LOCALE//./\\.}[[:space:]]+UTF-8)|\1|" /etc/locale.gen
else
  printf '%s UTF-8\n' "${TARGET_LOCALE}" >> /etc/locale.gen
fi

locale-gen "${TARGET_LOCALE}"

log "Setting system language to ${TARGET_LOCALE}..."
if command -v localectl >/dev/null 2>&1; then
  localectl set-locale "LANG=${TARGET_LOCALE}" "LANGUAGE=zh_CN:zh"
else
  update-locale "LANG=${TARGET_LOCALE}" "LANGUAGE=zh_CN:zh"
fi

if [[ ${SET_TIMEZONE} -eq 1 ]]; then
  log "Setting timezone to Asia/Shanghai..."
  timedatectl set-timezone Asia/Shanghai
fi

if [[ ${INSTALL_INPUT_METHOD} -eq 1 ]]; then
  TARGET_USER="${SUDO_USER:-}"
  if [[ -n "${TARGET_USER}" && "${TARGET_USER}" != "root" ]] && command -v runuser >/dev/null 2>&1; then
    log "Setting Fcitx5 as the default input method for user ${TARGET_USER}..."
    runuser -u "${TARGET_USER}" -- im-config -n fcitx5 || warn "im-config failed. Select Fcitx5 manually after logging in."
  else
    warn "No regular login user was detected. After logging in, run: im-config -n fcitx5"
  fi
fi

log "Done. Log out or reboot, then confirm the language is Chinese and the input method is Fcitx5."
