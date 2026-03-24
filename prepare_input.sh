#!/usr/bin/env bash

set -euo pipefail

usage() {
  echo "Usage: $(basename "$0") /path/to/archive.zip" >&2
}

if [[ $# -ne 1 ]]; then
  usage
  exit 1
fi

archive_path=$1

if [[ ! -f "$archive_path" ]]; then
  echo "Archive not found: $archive_path" >&2
  exit 1
fi

if [[ ${archive_path##*.} != "zip" && ${archive_path##*.} != "ZIP" ]]; then
  echo "Input file must have a .zip extension: $archive_path" >&2
  exit 1
fi

if ! command -v 7zz >/dev/null 2>&1; then
  echo "7zz is not installed or not available in PATH." >&2
  exit 1
fi

archive_dir=$(
  cd "$(dirname "$archive_path")"
  pwd
)
archive_name=$(basename "$archive_path")
input_dir="$archive_dir/input"

normalize_filename() {
  printf '%s' "$1" | tr -d '[:space:]_,-'
}

should_remove_report() {
  local normalized_name
  normalized_name=$(normalize_filename "$(basename "$1")")

  case "$normalized_name" in
    *Аттестацияэкстернов.xlsx|*Сведенияоборганизацииреализуемыхпрограммахиперсонале.xlsx)
      return 0
      ;;
  esac

  return 1
}

if [[ -e "$input_dir" && ! -d "$input_dir" ]]; then
  echo "Target path exists and is not a directory: $input_dir" >&2
  exit 1
fi

# Extract into a temporary location so unwanted files disappear with cleanup.
extract_dir=$(mktemp -d "${TMPDIR:-/tmp}/prepare-input.XXXXXX")
staging_dir=$(mktemp -d "$archive_dir/.input-staging.XXXXXX")
cleanup() {
  rm -rf "$extract_dir"
  rm -rf "$staging_dir"
}
trap cleanup EXIT INT TERM

echo "Extracting $archive_name"
7zz x -y "-o$extract_dir" "$archive_path" >/dev/null

svody_dirs=()
while IFS= read -r -d '' dir; do
  svody_dirs+=("$dir")
done < <(find "$extract_dir" -mindepth 1 -maxdepth 1 -type d -name 'Своды*' -print0)

if [[ ${#svody_dirs[@]} -eq 0 ]]; then
  echo "No top-level directory matching Своды* was found in the archive." >&2
  exit 1
fi

items_to_move=()
while IFS= read -r -d '' entry; do
  items_to_move+=("$entry")
done < <(
  for dir in "${svody_dirs[@]}"; do
    find "$dir" -mindepth 1 -maxdepth 1 ! -name '.DS_Store' -print0
  done
)

if [[ ${#items_to_move[@]} -eq 0 ]]; then
  echo "No content was found inside the matched Своды* directory." >&2
  exit 1
fi

for entry in "${items_to_move[@]}"; do
  target_path="$staging_dir/$(basename "$entry")"
  if [[ -e "$target_path" ]]; then
    echo "Refusing to overwrite existing path: $target_path" >&2
    exit 1
  fi
done

for entry in "${items_to_move[@]}"; do
  mv "$entry" "$staging_dir/"
done

removed_reports=0
while IFS= read -r -d '' file; do
  if should_remove_report "$file"; then
    rm -f "$file"
    removed_reports=$((removed_reports + 1))
  fi
done < <(find "$staging_dir" -type f -iname '*.xlsx' -print0)

rm -rf "$input_dir"
mv "$staging_dir" "$input_dir"
staging_dir=""

source_label="directories"
if [[ ${#svody_dirs[@]} -eq 1 ]]; then
  source_label="directory"
fi

item_label="items"
if [[ ${#items_to_move[@]} -eq 1 ]]; then
  item_label="item"
fi

echo "Prepared $input_dir with ${#items_to_move[@]} $item_label from ${#svody_dirs[@]} matching $source_label and removed $removed_reports report files"
