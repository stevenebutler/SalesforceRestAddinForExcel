#!/usr/bin/env bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

SOLUTION="${FC_SOLUTION:-SalesforceRestAddinForExcel.sln}"

run_restore() {
  dotnet restore "$SOLUTION" "$@"
}

run_build() {
  # net48 only: Core (Excel TFM), Windows.Ui, ExcelDna pack. Tests/net10 are for ./dev.sh test.
  local configuration="${FC_CONFIGURATION:-Debug}"
  echo "Building Core (net48, $configuration)..."
  dotnet build "$ROOT_DIR/src/SalesforceRestAddin.Core/SalesforceRestAddin.Core.csproj" \
    -c "$configuration" \
    -f net48 \
    --verbosity minimal \
    "$@"

  echo "Building SalesforceRestAddin.Windows.Ui (net48, $configuration)..."
  dotnet build "$ROOT_DIR/src/SalesforceRestAddin.Windows.Ui/SalesforceRestAddin.Windows.Ui.csproj" \
    -c "$configuration" \
    -f net48 \
    --verbosity minimal \
    "$@"

  if [[ "${FC_SKIP_EXCELDNA_BUILD:-0}" == "1" ]]; then
    echo "Skipping SalesforceRestAddin.ExcelDna build (FC_SKIP_EXCELDNA_BUILD=1)."
    return 0
  fi

  echo "Building SalesforceRestAddin.ExcelDna (net48, $configuration; packs XLL)..."
  if ! dotnet build "$ROOT_DIR/src/SalesforceRestAddin.ExcelDna/SalesforceRestAddin.ExcelDna.csproj" \
    -c "$configuration" \
    -f net48 \
    --verbosity minimal \
    "$@"; then
    echo "Build failed. On Linux, ExcelDna/WPF compile issues: see AGENTS.md (WPF refs + packing)." >&2
    return 1
  fi
}

run_test() {
  # net10.0 only: Core test TFM + TUnit (Linux runtime).
  echo "Building Core (net10.0) + Tests..."
  dotnet build "$ROOT_DIR/tests/SalesforceRestAddin.Tests/SalesforceRestAddin.Tests.csproj" \
    -f net10.0 \
    --verbosity minimal

  dotnet run --project "$ROOT_DIR/tests/SalesforceRestAddin.Tests/SalesforceRestAddin.Tests.csproj" \
    -f net10.0 \
    --no-build \
    -- \
    "$@"
}

COMMAND="${1:-shell}"

case "$COMMAND" in
  restore)
    shift || true
    run_restore "$@"
    ;;
  build)
    shift || true
    run_restore
    run_build "$@"
    ;;
  test)
    shift || true
    run_restore
    run_test "$@"
    ;;
  shell|bash|sh)
    shift || true
    if [[ $# -gt 0 ]]; then
      exec "$@"
    fi
    exec bash
    ;;
  down)
    echo "No long-running services to stop (dev container is ephemeral)."
    ;;
  *)
    exec "$COMMAND" "$@"
    ;;
esac
