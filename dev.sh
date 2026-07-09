#!/usr/bin/env bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

export COMPOSE_PROJECT_NAME="${COMPOSE_PROJECT_NAME:-salesforcerestaddin}"
export PODMAN_COMPOSE_WARNING_LOGS="${PODMAN_COMPOSE_WARNING_LOGS:-false}"

print_usage() {
  cat <<'EOF'
Usage: ./dev.sh <command> [args...]

Commands:
  build       Build net48 add-in path (Core + Windows.Ui + ExcelDna); copy XLLs to ~/Downloads/SalesforceRestAddin
  test        Build net10.0 Core + Tests; run TUnit
  restore     dotnet restore SalesforceRestAddinForExcel.sln
  shell       Interactive shell in the dev container
  down        Stop compose services
  help        Show this help

Any other command is passed through to the dev container entrypoint.

Examples:
  ./dev.sh build
  FC_CONFIGURATION=Release ./dev.sh build
  ./dev.sh test
  ./dev.sh test -- --filter SessionContext
  ./dev.sh bash -lc "dotnet --info"

Linux host convenience:
  alias dev=./dev.sh
EOF
}

run_compose() {
  docker compose "$@"
}

run_dev() {
  if [[ "${1:-}" == "shell" ]]; then
    run_compose run --rm dev "$@"
  else
    export FC_GIT_COMMIT="${FC_GIT_COMMIT:-$(git -C "$ROOT_DIR" describe --always --dirty 2>/dev/null || echo unknown)}"
    export FC_CONFIGURATION="${FC_CONFIGURATION:-Debug}"
    run_compose run -T --rm dev "$@"
  fi
}

copy_xll_to_downloads() {
  if [[ "${FC_SKIP_EXCELDNA_BUILD:-0}" == "1" ]]; then
    return 0
  fi
  copy_packed_xlls
}

copy_packed_xlls() {
  local configuration="${FC_CONFIGURATION:-Debug}"
  local publish_dir="$ROOT_DIR/src/SalesforceRestAddin.ExcelDna/bin/$configuration/net48/publish"
  local downloads_dir="${FC_DOWNLOADS_DIR:-$HOME/Downloads}"
  local deploy_dir="${FC_VM_DEPLOY_DIR:-$downloads_dir/SalesforceRestAddin}"

  if [[ ! -d "$publish_dir" ]]; then
    echo "No XLL publish directory at $publish_dir; skipping copy." >&2
    return 0
  fi

  mkdir -p "$deploy_dir"

  local copied=0
  local src dest base
  for src in "$publish_dir"/*-packed.xll; do
    [[ -f "$src" ]] || continue
    base="$(basename "$src")"
    if [[ "$base" == *AddIn64* ]]; then
      dest="$deploy_dir/SalesforceRestAddin64-packed.xll"
    else
      dest="$deploy_dir/SalesforceRestAddin-packed.xll"
    fi
    cp -f "$src" "$dest"
    echo "Copied $base -> $dest"
    copied=$((copied + 1))
  done

  if [[ "$copied" -eq 0 ]]; then
    echo "No packed XLL files found in $publish_dir" >&2
    return 1
  fi

  echo "Windows VM deploy bundle ready at $deploy_dir/"
}

COMMAND="${1:-}"

case "$COMMAND" in
  help|-h|--help)
    print_usage
    ;;
  shell)
    shift || true
    run_dev "$@"
    ;;
  build|test|restore|down)
    run_dev "$COMMAND" "${@:2}"
    if [[ "$COMMAND" == "build" ]]; then
      copy_xll_to_downloads
    fi
    ;;
  "")
    print_usage
    echo
    echo "Opening an interactive shell in the dev container..."
    echo
    run_dev shell
    ;;
  *)
    run_dev "$@"
    ;;
esac
