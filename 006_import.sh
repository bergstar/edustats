#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
sql_dir="$script_dir/006_output"

db_host="localhost"
db_user=""
db_password=""
db_name=""

usage() {
  cat <<'EOF'
Usage:
  ./006_import.sh --database NAME --user USER --password PASS [--host HOST]
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --database)
      [[ $# -ge 2 ]] || { usage >&2; exit 1; }
      db_name="$2"
      shift 2
      ;;
    --user)
      [[ $# -ge 2 ]] || { usage >&2; exit 1; }
      db_user="$2"
      shift 2
      ;;
    --password)
      [[ $# -ge 2 ]] || { usage >&2; exit 1; }
      db_password="$2"
      shift 2
      ;;
    --host)
      [[ $# -ge 2 ]] || { usage >&2; exit 1; }
      db_host="$2"
      shift 2
      ;;
    --help|-h)
      usage
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage >&2
      exit 1
      ;;
  esac
done

if [[ -z "$db_name" || -z "$db_user" || -z "$db_password" ]]; then
  usage >&2
  exit 1
fi

if [[ ! -d "$sql_dir" ]]; then
  echo "SQL directory does not exist: $sql_dir" >&2
  exit 1
fi

files=()

if [[ -f "$sql_dir/regions.sql" ]]; then
  files+=("$sql_dir/regions.sql")
fi

while IFS= read -r file; do
  [[ "$file" == "$sql_dir/regions.sql" ]] && continue
  files+=("$file")
done < <(find "$sql_dir" -maxdepth 1 -type f -name '*.sql' | sort)

if [[ ${#files[@]} -eq 0 ]]; then
  echo "No SQL files found in $sql_dir" >&2
  exit 1
fi

defaults_file="$(mktemp "${TMPDIR:-/tmp}/mysql-import.XXXXXX.cnf")"
cleanup() {
  rm -f "$defaults_file"
}
trap cleanup EXIT

chmod 600 "$defaults_file"
cat > "$defaults_file" <<EOF
[client]
protocol=tcp
host=$db_host
user=$db_user
password=$db_password
database=$db_name
default-character-set=utf8mb4
EOF

total_files="${#files[@]}"
current_file=0

drop_all_tables() {
  local tables_file
  tables_file="$(mktemp "${TMPDIR:-/tmp}/mysql-tables.XXXXXX.sql")"

  mysql --defaults-extra-file="$defaults_file" --batch --skip-column-names <<'SQL' > "$tables_file"
SELECT CONCAT('DROP TABLE IF EXISTS `', REPLACE(table_name, '`', '``'), '`;')
FROM information_schema.tables
WHERE table_schema = DATABASE()
ORDER BY table_name;
SQL

  if [[ -s "$tables_file" ]]; then
    {
      printf 'SET FOREIGN_KEY_CHECKS = 0;\n'
      cat "$tables_file"
      printf 'SET FOREIGN_KEY_CHECKS = 1;\n'
    } | mysql --defaults-extra-file="$defaults_file"
  fi

  rm -f "$tables_file"
}

printf 'Dropping existing tables in %s@%s/%s\n' "$db_user" "$db_host" "$db_name"
drop_all_tables

for file in "${files[@]}"; do
  current_file=$((current_file + 1))
  printf '\r\033[2K[%d/%d] %s' "$current_file" "$total_files" "$(basename "$file")"
  if ! mysql --defaults-extra-file="$defaults_file" < "$file"; then
    printf '\nFailed while importing %s\n' "$(basename "$file")" >&2
    exit 1
  fi
done

printf '\nImport completed. Imported %d files into %s@%s/%s\n' \
  "$total_files" \
  "$db_user" \
  "$db_host" \
  "$db_name"
