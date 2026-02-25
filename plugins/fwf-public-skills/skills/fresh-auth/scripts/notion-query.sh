#!/bin/bash
# Notion API CLI - Direct API access using curl
# Uses NOTION_API_KEY environment variable
# Optimized for token efficiency with jq filtering

set -e

# Load .env from skill root if it exists
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
if [ -f "$SKILL_ROOT/.env" ]; then
    set -a
    source "$SKILL_ROOT/.env"
    set +a
fi

# Default to legacy API - set NOTION_API_VERSION=2025-09-03 for new data source support
NOTION_VERSION="${NOTION_API_VERSION:-2022-06-28}"
API_BASE="https://api.notion.com/v1"
BACKLOG_DB_ID="${NOTION_BACKLOG_DB_ID:-}"

# Check for API key
if [ -z "$NOTION_API_KEY" ]; then
    echo "Error: NOTION_API_KEY environment variable not set" >&2
    exit 1
fi

# Helper function for API calls
notion_api() {
    local method="$1"
    local endpoint="$2"
    local data="$3"

    if [ -n "$data" ]; then
        curl -s -X "$method" "$API_BASE$endpoint" \
            -H "Authorization: Bearer $NOTION_API_KEY" \
            -H "Notion-Version: $NOTION_VERSION" \
            -H "Content-Type: application/json" \
            -d "$data"
    else
        curl -s -X "$method" "$API_BASE$endpoint" \
            -H "Authorization: Bearer $NOTION_API_KEY" \
            -H "Notion-Version: $NOTION_VERSION"
    fi
}

# Check for API errors and exit if found
check_error() {
    local result="$1"
    if echo "$result" | jq -e '.object == "error"' > /dev/null 2>&1; then
        echo "$result" | jq -r '"Error: \(.message)"' >&2
        exit 1
    fi
}

# Check if a value looks like a Notion object ID (32 chars or UUID style)
is_probable_notion_id() {
    local value="$1"
    [[ "$value" =~ ^[0-9a-fA-F-]{32,36}$ ]]
}

# Discover databases to help users find DB IDs
find_databases() {
    local query="${1:-}"
    local payload
    payload=$(jq -n --arg q "$query" '{
        query: $q,
        page_size: 100,
        filter: {property: "object", value: "database"}
    }')

    local result
    result=$(notion_api POST "/search" "$payload")
    check_error "$result"

    echo "$result" | jq '[.results[] | {
        id: .id,
        title: (.title[0].plain_text // "Untitled"),
        url: .url
    }]'
}

print_find_db_hint() {
    echo "Need a database ID? Discover it with:" >&2
    echo "  notion-query.sh find-db \"keyword\"" >&2
    echo "  notion-query.sh find-db" >&2
}

# Extract page title from properties
extract_title() {
    jq -r '
        if .properties.Name.title[0].plain_text then
            .properties.Name.title[0].plain_text
        elif .properties.title.title[0].plain_text then
            .properties.title.title[0].plain_text
        else
            "Untitled"
        end
    '
}

# Helper function to convert markdown text to Notion rich_text JSON array
# Supports **bold**, *italic*, `code`, ~~strikethrough~~
markdown_to_rich_text() {
    local text="$1"
    local result="["
    local first=true
    local current=""
    local i=0
    local len=${#text}

    while [ $i -lt $len ]; do
        local char="${text:$i:1}"
        local next="${text:$((i+1)):1}"

        # Check for **bold**
        if [ "$char" = "*" ] && [ "$next" = "*" ]; then
            # Output current plain text if any
            if [ -n "$current" ]; then
                [ "$first" = true ] || result+=","
                first=false
                current="${current//\\/\\\\}"
                current="${current//\"/\\\"}"
                result+="{\"type\":\"text\",\"text\":{\"content\":\"$current\"}}"
                current=""
            fi
            # Find closing **
            local j=$((i+2))
            local bold_text=""
            while [ $j -lt $((len-1)) ]; do
                if [ "${text:$j:2}" = "**" ]; then
                    break
                fi
                bold_text+="${text:$j:1}"
                j=$((j+1))
            done
            if [ "${text:$j:2}" = "**" ]; then
                [ "$first" = true ] || result+=","
                first=false
                bold_text="${bold_text//\\/\\\\}"
                bold_text="${bold_text//\"/\\\"}"
                result+="{\"type\":\"text\",\"text\":{\"content\":\"$bold_text\"},\"annotations\":{\"bold\":true}}"
                i=$((j+2))
                continue
            fi
        fi

        # Check for ~~strikethrough~~
        if [ "$char" = "~" ] && [ "$next" = "~" ]; then
            if [ -n "$current" ]; then
                [ "$first" = true ] || result+=","
                first=false
                current="${current//\\/\\\\}"
                current="${current//\"/\\\"}"
                result+="{\"type\":\"text\",\"text\":{\"content\":\"$current\"}}"
                current=""
            fi
            local j=$((i+2))
            local strike_text=""
            while [ $j -lt $((len-1)) ]; do
                if [ "${text:$j:2}" = "~~" ]; then
                    break
                fi
                strike_text+="${text:$j:1}"
                j=$((j+1))
            done
            if [ "${text:$j:2}" = "~~" ]; then
                [ "$first" = true ] || result+=","
                first=false
                strike_text="${strike_text//\\/\\\\}"
                strike_text="${strike_text//\"/\\\"}"
                result+="{\"type\":\"text\",\"text\":{\"content\":\"$strike_text\"},\"annotations\":{\"strikethrough\":true}}"
                i=$((j+2))
                continue
            fi
        fi

        # Check for `code`
        if [ "$char" = "\`" ] && [ "$next" != "\`" ]; then
            if [ -n "$current" ]; then
                [ "$first" = true ] || result+=","
                first=false
                current="${current//\\/\\\\}"
                current="${current//\"/\\\"}"
                result+="{\"type\":\"text\",\"text\":{\"content\":\"$current\"}}"
                current=""
            fi
            local j=$((i+1))
            local code_text=""
            while [ $j -lt $len ]; do
                if [ "${text:$j:1}" = "\`" ]; then
                    break
                fi
                code_text+="${text:$j:1}"
                j=$((j+1))
            done
            if [ "${text:$j:1}" = "\`" ]; then
                [ "$first" = true ] || result+=","
                first=false
                code_text="${code_text//\\/\\\\}"
                code_text="${code_text//\"/\\\"}"
                result+="{\"type\":\"text\",\"text\":{\"content\":\"$code_text\"},\"annotations\":{\"code\":true}}"
                i=$((j+1))
                continue
            fi
        fi

        # Check for *italic* (single asterisk, not followed by another)
        if [ "$char" = "*" ] && [ "$next" != "*" ]; then
            if [ -n "$current" ]; then
                [ "$first" = true ] || result+=","
                first=false
                current="${current//\\/\\\\}"
                current="${current//\"/\\\"}"
                result+="{\"type\":\"text\",\"text\":{\"content\":\"$current\"}}"
                current=""
            fi
            local j=$((i+1))
            local italic_text=""
            while [ $j -lt $len ]; do
                if [ "${text:$j:1}" = "*" ] && [ "${text:$((j+1)):1}" != "*" ]; then
                    break
                fi
                italic_text+="${text:$j:1}"
                j=$((j+1))
            done
            if [ "${text:$j:1}" = "*" ]; then
                [ "$first" = true ] || result+=","
                first=false
                italic_text="${italic_text//\\/\\\\}"
                italic_text="${italic_text//\"/\\\"}"
                result+="{\"type\":\"text\",\"text\":{\"content\":\"$italic_text\"},\"annotations\":{\"italic\":true}}"
                i=$((j+1))
                continue
            fi
        fi

        # Regular character
        current+="$char"
        i=$((i+1))
    done

    # Output remaining text
    if [ -n "$current" ]; then
        [ "$first" = true ] || result+=","
        current="${current//\\/\\\\}"
        current="${current//\"/\\\"}"
        result+="{\"type\":\"text\",\"text\":{\"content\":\"$current\"}}"
    fi

    result+="]"
    echo "$result"
}

# Commands
case "$1" in
    search)
        # Search pages and databases - returns compact results
        # Usage: notion-query.sh search [query]
        query="${2:-}"
        payload=$(jq -n --arg q "$query" '{query: $q, page_size: 100}')
        result=$(notion_api POST "/search" "$payload")
        check_error "$result"
        echo "$result" | jq '[.results[] | {
            id: .id,
            type: .object,
            title: (
                if .object == "database" then
                    .title[0].plain_text // "Untitled"
                else
                    (.properties.Name.title[0].plain_text // .properties.title.title[0].plain_text // "Untitled")
                end
            ),
            url: .url
        }]'
        ;;

    find-db|find-dbs|list-dbs)
        # Search databases only - helpful when user does not know database ID
        # Usage: notion-query.sh find-db [query]
        query="${2:-}"
        find_databases "$query"
        ;;

    query-db)
        # Query a database - returns compact results with all properties
        # Usage: notion-query.sh query-db <database_id> [--raw] [--filter-prop=value]
        db_id="$2"
        raw_mode=false
        filter_prop=""
        filter_val=""

        if [ -z "$db_id" ]; then
            echo "Usage: notion-query.sh query-db <database_id> [--raw] [--filter-prop=value]" >&2
            echo "" >&2
            print_find_db_hint
            exit 1
        fi
        if ! is_probable_notion_id "$db_id"; then
            echo "Error: '$db_id' does not look like a Notion database ID." >&2
            echo "Database matches for '$db_id':" >&2
            find_databases "$db_id"
            exit 1
        fi

        shift 2
        while [ $# -gt 0 ]; do
            case "$1" in
                --raw)
                    raw_mode=true
                    ;;
                --filter-*)
                    filter_prop="${1#--filter-}"
                    filter_prop="${filter_prop%%=*}"
                    filter_val="${1#*=}"
                    ;;
            esac
            shift
        done

        result=$(notion_api POST "/databases/$db_id/query" "{}")
        check_error "$result"

        if [ "$raw_mode" = true ]; then
            echo "$result"
        else
            # Extract compact view with all properties dynamically
            echo "$result" | jq '[.results[] | {
                id: .id,
                url: .url
            } + (.properties | to_entries | map({
                key: .key,
                value: (
                    if .value.title then (.value.title | map(.plain_text) | join(""))
                    elif .value.rich_text then (.value.rich_text | map(.plain_text) | join(""))
                    elif .value.select then .value.select.name
                    elif .value.multi_select then (.value.multi_select | map(.name) | join(", "))
                    elif .value.status then .value.status.name
                    elif .value.number then .value.number
                    elif .value.checkbox then .value.checkbox
                    elif .value.date then .value.date.start
                    elif .value.url then .value.url
                    elif .value.email then .value.email
                    elif .value.phone_number then .value.phone_number
                    elif .value.people then (.value.people | map(.name // .id) | join(", "))
                    elif .value.relation then (.value.relation | map(.id) | join(", "))
                    elif .value.formula then (
                        if .value.formula.string then .value.formula.string
                        elif .value.formula.number then .value.formula.number
                        elif .value.formula.boolean then .value.formula.boolean
                        elif .value.formula.date then .value.formula.date.start
                        else null
                        end
                    )
                    elif .value.rollup then (
                        if .value.rollup.array then (.value.rollup.array | length | tostring) + " items"
                        elif .value.rollup.number then .value.rollup.number
                        else null
                        end
                    )
                    elif .value.created_time then .value.created_time
                    elif .value.last_edited_time then .value.last_edited_time
                    elif .value.created_by then .value.created_by.name
                    elif .value.last_edited_by then .value.last_edited_by.name
                    elif .value.files then (.value.files | map(.name // .file.url // .external.url) | join(", "))
                    else null
                    end
                )
            }) | from_entries)]' | if [ -n "$filter_prop" ] && [ -n "$filter_val" ]; then
                jq --arg prop "$filter_prop" --arg val "$filter_val" '[.[] | select(.[$prop] == $val)]'
            else
                cat
            fi
        fi
        ;;

    backlog)
        # Query configured backlog database with optional roadmap filter
        # Usage: notion-query.sh backlog [roadmap_filter]
        # Examples: notion-query.sh backlog "In Progress"
        #           notion-query.sh backlog "Top 3"
        if [ -z "$BACKLOG_DB_ID" ]; then
            echo "Error: NOTION_BACKLOG_DB_ID environment variable not set" >&2
            echo "Set NOTION_BACKLOG_DB_ID to use backlog shortcuts." >&2
            exit 1
        fi
        db_id="$BACKLOG_DB_ID"
        roadmap_filter="${2:-}"

        result=$(notion_api POST "/databases/$db_id/query" "{}")
        check_error "$result"

        echo "$result" | jq --arg filter "$roadmap_filter" '
            [.results[] | {
                title: .properties.Name.title[0].plain_text,
                roadmap: .properties.Roadmap.select.name,
                status: .properties.Status.status.name
            }] |
            if $filter != "" then
                [.[] | select(.roadmap == $filter)]
            else
                .
            end |
            group_by(.roadmap) |
            map({
                roadmap: .[0].roadmap,
                items: [.[] | {title, status}]
            })
        '
        ;;

    get-page)
        # Get a page - returns compact properties
        # Usage: notion-query.sh get-page <page_id> [--raw]
        page_id="$2"
        raw_mode=false
        [ "$3" = "--raw" ] && raw_mode=true

        if [ -z "$page_id" ]; then
            echo "Usage: notion-query.sh get-page <page_id> [--raw]" >&2
            exit 1
        fi

        result=$(notion_api GET "/pages/$page_id")
        check_error "$result"

        if [ "$raw_mode" = true ]; then
            echo "$result"
        else
            echo "$result" | jq '{
                id: .id,
                title: (.properties.Name.title[0].plain_text // .properties.title.title[0].plain_text // "Untitled"),
                properties: (.properties | to_entries | map({
                    key: .key,
                    value: (
                        if .value.select then .value.select.name
                        elif .value.status then .value.status.name
                        elif .value.title then .value.title[0].plain_text
                        elif .value.rich_text then (.value.rich_text | map(.plain_text) | join(""))
                        elif .value.number then .value.number
                        elif .value.checkbox then .value.checkbox
                        elif .value.date then .value.date.start
                        else null
                        end
                    )
                }) | from_entries),
                url: .url
            }'
        fi
        ;;

    get-blocks)
        # Get block children - returns text content only
        # Usage: notion-query.sh get-blocks <block_id> [--raw]
        block_id="$2"
        raw_mode=false
        [ "$3" = "--raw" ] && raw_mode=true

        if [ -z "$block_id" ]; then
            echo "Usage: notion-query.sh get-blocks <block_id> [--raw]" >&2
            exit 1
        fi

        result=$(notion_api GET "/blocks/$block_id/children?page_size=100")
        check_error "$result"

        if [ "$raw_mode" = true ]; then
            echo "$result"
        else
            echo "$result" | jq '[.results[] | {
                type: .type,
                id: .id,
                text: (
                    if .paragraph then (.paragraph.rich_text | map(.plain_text) | join(""))
                    elif .heading_1 then (.heading_1.rich_text | map(.plain_text) | join(""))
                    elif .heading_2 then (.heading_2.rich_text | map(.plain_text) | join(""))
                    elif .heading_3 then (.heading_3.rich_text | map(.plain_text) | join(""))
                    elif .bulleted_list_item then (.bulleted_list_item.rich_text | map(.plain_text) | join(""))
                    elif .numbered_list_item then (.numbered_list_item.rich_text | map(.plain_text) | join(""))
                    elif .to_do then (.to_do.rich_text | map(.plain_text) | join(""))
                    elif .code then (.code.rich_text | map(.plain_text) | join(""))
                    elif .quote then (.quote.rich_text | map(.plain_text) | join(""))
                    elif .callout then (.callout.rich_text | map(.plain_text) | join(""))
                    elif .toggle then (.toggle.rich_text | map(.plain_text) | join(""))
                    elif .child_page then .child_page.title
                    elif .child_database then .child_database.title
                    else null
                    end
                ),
                checked: (if .to_do then .to_do.checked else null end)
            } | select(.text != null and .text != "")]'
        fi
        ;;

    get-db)
        # Get database schema - returns property names and types only
        # Usage: notion-query.sh get-db <database_id> [--raw]
        db_id="$2"
        raw_mode=false
        [ "$3" = "--raw" ] && raw_mode=true

        if [ -z "$db_id" ]; then
            echo "Usage: notion-query.sh get-db <database_id> [--raw]" >&2
            echo "" >&2
            print_find_db_hint
            exit 1
        fi
        if ! is_probable_notion_id "$db_id"; then
            echo "Error: '$db_id' does not look like a Notion database ID." >&2
            echo "Database matches for '$db_id':" >&2
            find_databases "$db_id"
            exit 1
        fi

        result=$(notion_api GET "/databases/$db_id")
        check_error "$result"

        if [ "$raw_mode" = true ]; then
            echo "$result"
        else
            echo "$result" | jq '{
                id: .id,
                title: .title[0].plain_text,
                properties: (.properties | to_entries | map({
                    name: .key,
                    type: .value.type,
                    options: (
                        if .value.select.options then [.value.select.options[].name]
                        elif .value.status.options then [.value.status.options[].name]
                        elif .value.multi_select.options then [.value.multi_select.options[].name]
                        else null
                        end
                    )
                }))
            }'
        fi
        ;;

    create)
        # Create a new page in a database
        # Usage: notion-query.sh create <db_id> <title> [-p NAME=VALUE]...
        # Examples:
        #   notion-query.sh create <db_id> "Title" -p "Status=In progress" -p "Priority=High"
        #   notion-query.sh create <db_id> "Title" --select-Roadmap="Soon" --status-Status="Not started"
        db_id="$2"
        title="$3"
        props_json="{}"

        if [ -z "$db_id" ] || [ -z "$title" ]; then
            echo "Usage: notion-query.sh create <db_id> <title> [-p NAME=VALUE]..." >&2
            echo "       notion-query.sh create <db_id> <title> [--TYPE-NAME=VALUE]..." >&2
            echo "" >&2
            echo "Property types: --select-NAME, --status-NAME, --text-NAME, --number-NAME, --checkbox-NAME, --date-NAME, --url-NAME" >&2
            echo "" >&2
            print_find_db_hint
            exit 1
        fi
        if ! is_probable_notion_id "$db_id"; then
            echo "Error: '$db_id' does not look like a Notion database ID." >&2
            echo "Database matches for '$db_id':" >&2
            find_databases "$db_id"
            exit 1
        fi

        # First, get database schema to auto-detect property types
        db_schema=$(notion_api GET "/databases/$db_id")
        check_error "$db_schema"

        shift 3
        while [ $# -gt 0 ]; do
            case "$1" in
                -p)
                    # Generic property: -p "Name=Value" - auto-detect type from schema
                    shift
                    if [ -n "$1" ]; then
                        prop_name="${1%%=*}"
                        prop_val="${1#*=}"
                        # Get property type from schema
                        prop_type=$(echo "$db_schema" | jq -r --arg name "$prop_name" '.properties[$name].type // "unknown"')
                        case "$prop_type" in
                            select)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {select: {name: $val}}}')
                                ;;
                            status)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {status: {name: $val}}}')
                                ;;
                            multi_select)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {multi_select: [{name: $val}]}}')
                                ;;
                            rich_text)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {rich_text: [{text: {content: $val}}]}}')
                                ;;
                            number)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$prop_val" '. + {($name): {number: $val}}')
                                ;;
                            checkbox)
                                bool_val="false"
                                [ "$prop_val" = "true" ] || [ "$prop_val" = "1" ] || [ "$prop_val" = "yes" ] && bool_val="true"
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$bool_val" '. + {($name): {checkbox: $val}}')
                                ;;
                            date)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {date: {start: $val}}}')
                                ;;
                            url)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {url: $val}}')
                                ;;
                            email)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {email: $val}}')
                                ;;
                            phone_number)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {phone_number: $val}}')
                                ;;
                            relation)
                                # Relation expects page ID(s) - supports comma-separated IDs
                                IFS=',' read -ra ids <<< "$prop_val"
                                relation_array="[]"
                                for id in "${ids[@]}"; do
                                    relation_array=$(echo "$relation_array" | jq --arg id "$id" '. + [{id: $id}]')
                                done
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$relation_array" '. + {($name): {relation: $val}}')
                                ;;
                            *)
                                echo "Warning: Unknown property type '$prop_type' for '$prop_name', trying as select" >&2
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {select: {name: $val}}}')
                                ;;
                        esac
                    fi
                    ;;
                --select-*)
                    prop_name="${1#--select-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {select: {name: $val}}}')
                    ;;
                --status-*)
                    prop_name="${1#--status-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {status: {name: $val}}}')
                    ;;
                --text-*)
                    prop_name="${1#--text-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {rich_text: [{text: {content: $val}}]}}')
                    ;;
                --number-*)
                    prop_name="${1#--number-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$prop_val" '. + {($name): {number: $val}}')
                    ;;
                --checkbox-*)
                    prop_name="${1#--checkbox-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    bool_val="false"
                    [ "$prop_val" = "true" ] || [ "$prop_val" = "1" ] || [ "$prop_val" = "yes" ] && bool_val="true"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$bool_val" '. + {($name): {checkbox: $val}}')
                    ;;
                --date-*)
                    prop_name="${1#--date-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {date: {start: $val}}}')
                    ;;
                --url-*)
                    prop_name="${1#--url-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {url: $val}}')
                    ;;
                --multi-select-*)
                    prop_name="${1#--multi-select-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {multi_select: [{name: $val}]}}')
                    ;;
            esac
            shift
        done

        # Add title to properties
        props=$(echo "$props_json" | jq --arg title "$title" '. + {Name: {title: [{text: {content: $title}}]}}')

        payload=$(jq -n --arg db_id "$db_id" --argjson props "$props" \
            '{parent: {database_id: $db_id}, properties: $props}')

        result=$(notion_api POST "/pages" "$payload")
        check_error "$result"

        echo "$result" | jq '{
            id: .id,
            url: .url
        } + (.properties | to_entries | map(select(.value.title)) | if length > 0 then {title: .[0].value.title[0].plain_text} else {} end)'
        ;;

    create-backlog)
        # Shortcut to create item in configured backlog database
        # Usage: notion-query.sh create-backlog <title> [--roadmap=X] [--status=X]
        title="$2"
        if [ -z "$title" ]; then
            echo "Usage: notion-query.sh create-backlog <title> [--roadmap=X] [--status=X]" >&2
            exit 1
        fi
        if [ -z "$BACKLOG_DB_ID" ]; then
            echo "Error: NOTION_BACKLOG_DB_ID environment variable not set" >&2
            echo "Set NOTION_BACKLOG_DB_ID to use backlog shortcuts." >&2
            exit 1
        fi
        # Forward to create with backlog db_id
        shift 1
        exec "$0" create "$BACKLOG_DB_ID" "$@"
        ;;

    update)
        # Update page properties
        # Usage: notion-query.sh update <page_id> [-p NAME=VALUE]...
        # Examples:
        #   notion-query.sh update <page_id> -p "Status=Done" -p "Priority=Low"
        #   notion-query.sh update <page_id> --title="New Title" --select-Priority="High"
        page_id="$2"
        props_json="{}"
        has_props=false

        if [ -z "$page_id" ]; then
            echo "Usage: notion-query.sh update <page_id> [-p NAME=VALUE]..." >&2
            echo "       notion-query.sh update <page_id> [--title=X] [--TYPE-NAME=VALUE]..." >&2
            echo "" >&2
            echo "Property types: --select-NAME, --status-NAME, --text-NAME, --number-NAME, --checkbox-NAME, --date-NAME, --url-NAME" >&2
            exit 1
        fi

        # Get page to find its parent database for schema lookup
        page_info=$(notion_api GET "/pages/$page_id")
        check_error "$page_info"
        parent_db_id=$(echo "$page_info" | jq -r '.parent.database_id // empty')

        db_schema="{}"
        if [ -n "$parent_db_id" ]; then
            db_schema=$(notion_api GET "/databases/$parent_db_id" 2>/dev/null || echo "{}")
        fi

        shift 2
        while [ $# -gt 0 ]; do
            case "$1" in
                --title=*)
                    title="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg val "$title" '. + {Name: {title: [{text: {content: $val}}]}}')
                    has_props=true
                    ;;
                -p)
                    # Generic property: -p "Name=Value" - auto-detect type from schema
                    shift
                    if [ -n "$1" ]; then
                        prop_name="${1%%=*}"
                        prop_val="${1#*=}"
                        # Get property type from schema
                        prop_type=$(echo "$db_schema" | jq -r --arg name "$prop_name" '.properties[$name].type // "unknown"')
                        case "$prop_type" in
                            select)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {select: {name: $val}}}')
                                ;;
                            status)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {status: {name: $val}}}')
                                ;;
                            multi_select)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {multi_select: [{name: $val}]}}')
                                ;;
                            rich_text)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {rich_text: [{text: {content: $val}}]}}')
                                ;;
                            number)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$prop_val" '. + {($name): {number: $val}}')
                                ;;
                            checkbox)
                                bool_val="false"
                                [ "$prop_val" = "true" ] || [ "$prop_val" = "1" ] || [ "$prop_val" = "yes" ] && bool_val="true"
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$bool_val" '. + {($name): {checkbox: $val}}')
                                ;;
                            date)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {date: {start: $val}}}')
                                ;;
                            url)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {url: $val}}')
                                ;;
                            email)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {email: $val}}')
                                ;;
                            phone_number)
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {phone_number: $val}}')
                                ;;
                            relation)
                                # Relation expects page ID(s) - supports comma-separated IDs
                                IFS=',' read -ra ids <<< "$prop_val"
                                relation_array="[]"
                                for id in "${ids[@]}"; do
                                    relation_array=$(echo "$relation_array" | jq --arg id "$id" '. + [{id: $id}]')
                                done
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$relation_array" '. + {($name): {relation: $val}}')
                                ;;
                            *)
                                echo "Warning: Unknown property type '$prop_type' for '$prop_name', trying as select" >&2
                                props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {select: {name: $val}}}')
                                ;;
                        esac
                        has_props=true
                    fi
                    ;;
                --select-*)
                    prop_name="${1#--select-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {select: {name: $val}}}')
                    has_props=true
                    ;;
                --status-*)
                    prop_name="${1#--status-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {status: {name: $val}}}')
                    has_props=true
                    ;;
                --text-*)
                    prop_name="${1#--text-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {rich_text: [{text: {content: $val}}]}}')
                    has_props=true
                    ;;
                --number-*)
                    prop_name="${1#--number-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$prop_val" '. + {($name): {number: $val}}')
                    has_props=true
                    ;;
                --checkbox-*)
                    prop_name="${1#--checkbox-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    bool_val="false"
                    [ "$prop_val" = "true" ] || [ "$prop_val" = "1" ] || [ "$prop_val" = "yes" ] && bool_val="true"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --argjson val "$bool_val" '. + {($name): {checkbox: $val}}')
                    has_props=true
                    ;;
                --date-*)
                    prop_name="${1#--date-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {date: {start: $val}}}')
                    has_props=true
                    ;;
                --url-*)
                    prop_name="${1#--url-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {url: $val}}')
                    has_props=true
                    ;;
                --multi-select-*)
                    prop_name="${1#--multi-select-}"
                    prop_name="${prop_name%%=*}"
                    prop_val="${1#*=}"
                    props_json=$(echo "$props_json" | jq --arg name "$prop_name" --arg val "$prop_val" '. + {($name): {multi_select: [{name: $val}]}}')
                    has_props=true
                    ;;
            esac
            shift
        done

        if [ "$has_props" = false ]; then
            echo "Error: No properties to update. Use -p NAME=VALUE or --TYPE-NAME=VALUE" >&2
            exit 1
        fi

        payload=$(jq -n --argjson props "$props_json" '{properties: $props}')
        result=$(notion_api PATCH "/pages/$page_id" "$payload")
        check_error "$result"

        # Return compact view with all updated properties
        echo "$result" | jq '{
            id: .id,
            url: .url
        } + (.properties | to_entries | map({
            key: .key,
            value: (
                if .value.title then (.value.title | map(.plain_text) | join(""))
                elif .value.rich_text then (.value.rich_text | map(.plain_text) | join(""))
                elif .value.select then .value.select.name
                elif .value.multi_select then (.value.multi_select | map(.name) | join(", "))
                elif .value.status then .value.status.name
                elif .value.number then .value.number
                elif .value.checkbox then .value.checkbox
                elif .value.date then .value.date.start
                elif .value.url then .value.url
                else null
                end
            )
        }) | from_entries)'
        ;;

    archive)
        # Archive (trash) a page
        # Usage: notion-query.sh archive <page_id>
        page_id="$2"

        if [ -z "$page_id" ]; then
            echo "Usage: notion-query.sh archive <page_id>" >&2
            exit 1
        fi

        result=$(notion_api PATCH "/pages/$page_id" '{"archived": true}')
        check_error "$result"

        echo "$result" | jq '{
            id: .id,
            title: (.properties.Name.title[0].plain_text // .properties.title.title[0].plain_text // "Untitled"),
            archived: .archived
        }'
        ;;

    get-markdown)
        # Get page content as markdown
        # Usage: notion-query.sh get-markdown <page_id>
        page_id="$2"

        if [ -z "$page_id" ]; then
            echo "Usage: notion-query.sh get-markdown <page_id>" >&2
            exit 1
        fi

        result=$(notion_api GET "/blocks/$page_id/children?page_size=100")
        check_error "$result"

        echo "$result" | jq -r '.results[] |
            if .type == "heading_1" then
                "# " + (.heading_1.rich_text | map(
                    if .href then "[" + .plain_text + "](" + .href + ")"
                    else .plain_text end
                ) | join(""))
            elif .type == "heading_2" then
                "## " + (.heading_2.rich_text | map(
                    if .href then "[" + .plain_text + "](" + .href + ")"
                    else .plain_text end
                ) | join(""))
            elif .type == "heading_3" then
                "### " + (.heading_3.rich_text | map(
                    if .href then "[" + .plain_text + "](" + .href + ")"
                    else .plain_text end
                ) | join(""))
            elif .type == "paragraph" then
                (.paragraph.rich_text | map(
                    if .href then "[" + .plain_text + "](" + .href + ")"
                    else .plain_text end
                ) | join(""))
            elif .type == "bulleted_list_item" then
                "- " + (.bulleted_list_item.rich_text | map(
                    if .href then "[" + .plain_text + "](" + .href + ")"
                    else .plain_text end
                ) | join(""))
            elif .type == "numbered_list_item" then
                "1. " + (.numbered_list_item.rich_text | map(
                    if .href then "[" + .plain_text + "](" + .href + ")"
                    else .plain_text end
                ) | join(""))
            elif .type == "to_do" then
                (if .to_do.checked then "- [x] " else "- [ ] " end) + (.to_do.rich_text | map(
                    if .href then "[" + .plain_text + "](" + .href + ")"
                    else .plain_text end
                ) | join(""))
            elif .type == "quote" then
                "> " + (.quote.rich_text | map(
                    if .href then "[" + .plain_text + "](" + .href + ")"
                    else .plain_text end
                ) | join(""))
            elif .type == "code" then
                "```" + (.code.language // "") + "\n" + (.code.rich_text | map(.plain_text) | join("")) + "\n```"
            elif .type == "divider" then
                "---"
            elif .type == "child_page" then
                "📄 " + .child_page.title
            elif .type == "child_database" then
                "📊 " + .child_database.title
            elif .type == "callout" then
                "> " + (.callout.icon.emoji // "💡") + " " + (.callout.rich_text | map(.plain_text) | join(""))
            else
                empty
            end
        '
        ;;

    set-body)
        # Set page content from markdown (replaces existing content)
        # Usage: notion-query.sh set-body <page_id> <markdown>
        # Or:    echo "markdown" | notion-query.sh set-body <page_id> -
        page_id="$2"
        markdown_input="$3"

        if [ -z "$page_id" ]; then
            echo "Usage: notion-query.sh set-body <page_id> <markdown>" >&2
            echo "       echo 'markdown' | notion-query.sh set-body <page_id> -" >&2
            exit 1
        fi

        # Read from stdin if - is passed
        if [ "$markdown_input" = "-" ]; then
            markdown_input=$(cat)
        fi

        if [ -z "$markdown_input" ]; then
            echo "Error: No markdown content provided" >&2
            exit 1
        fi

        # First, delete existing blocks
        existing=$(notion_api GET "/blocks/$page_id/children?page_size=100")
        check_error "$existing"

        # Delete each existing block
        echo "$existing" | jq -r '.results[].id' | while read -r block_id; do
            if [ -n "$block_id" ]; then
                notion_api DELETE "/blocks/$block_id" > /dev/null 2>&1 || true
            fi
        done

        # Convert markdown to Notion blocks JSON
        blocks="["
        first_block=true
        in_code=false
        code_lang=""
        code_content=""

        while IFS= read -r line || [ -n "$line" ]; do
            # Code block handling
            if [[ "$line" =~ ^\`\`\` ]]; then
                if [ "$in_code" = false ]; then
                    in_code=true
                    code_lang="${line:3}"
                    code_content=""
                else
                    in_code=false
                    [ "$first_block" = true ] || blocks+=","
                    first_block=false
                    code_content="${code_content//\\/\\\\}"
                    code_content="${code_content//\"/\\\"}"
                    [ -z "$code_lang" ] && code_lang="plain text"
                    blocks+="{\"object\":\"block\",\"type\":\"code\",\"code\":{\"language\":\"$code_lang\",\"rich_text\":[{\"type\":\"text\",\"text\":{\"content\":\"$code_content\"}}]}}"
                fi
                continue
            fi

            if [ "$in_code" = true ]; then
                [ -n "$code_content" ] && code_content+="\\n"
                escaped_line="${line//\\/\\\\}"
                escaped_line="${escaped_line//\"/\\\"}"
                code_content+="$escaped_line"
                continue
            fi

            # Skip empty lines
            [[ "$line" =~ ^[[:space:]]*$ ]] && continue

            [ "$first_block" = true ] || blocks+=","
            first_block=false

            # Heading 1
            if [[ "$line" =~ ^#\  ]]; then
                text="${line:2}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"heading_1\",\"heading_1\":{\"rich_text\":$rich_text}}"
            # Heading 2
            elif [[ "$line" =~ ^##\  ]]; then
                text="${line:3}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"heading_2\",\"heading_2\":{\"rich_text\":$rich_text}}"
            # Heading 3
            elif [[ "$line" =~ ^###\  ]]; then
                text="${line:4}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"heading_3\",\"heading_3\":{\"rich_text\":$rich_text}}"
            # Divider
            elif [[ "$line" =~ ^---$ ]] || [[ "$line" =~ ^\*\*\*$ ]] || [[ "$line" =~ ^___$ ]]; then
                blocks+="{\"object\":\"block\",\"type\":\"divider\",\"divider\":{}}"
            # Todo checked
            elif [[ "$line" =~ ^-\ \[[xX]\]\  ]]; then
                text="${line:6}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"to_do\",\"to_do\":{\"checked\":true,\"rich_text\":$rich_text}}"
            # Todo unchecked
            elif [[ "$line" =~ ^-\ \[\ \]\  ]]; then
                text="${line:6}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"to_do\",\"to_do\":{\"checked\":false,\"rich_text\":$rich_text}}"
            # Bullet list
            elif [[ "$line" =~ ^-\  ]] || [[ "$line" =~ ^\*\  ]]; then
                text="${line:2}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"bulleted_list_item\",\"bulleted_list_item\":{\"rich_text\":$rich_text}}"
            # Numbered list
            elif [[ "$line" =~ ^[0-9]+\.\  ]]; then
                text="${line#*. }"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"numbered_list_item\",\"numbered_list_item\":{\"rich_text\":$rich_text}}"
            # Quote
            elif [[ "$line" =~ ^\>\  ]]; then
                text="${line:2}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"quote\",\"quote\":{\"rich_text\":$rich_text}}"
            # Regular paragraph
            else
                rich_text=$(markdown_to_rich_text "$line")
                blocks+="{\"object\":\"block\",\"type\":\"paragraph\",\"paragraph\":{\"rich_text\":$rich_text}}"
            fi
        done <<< "$markdown_input"

        blocks+="]"

        # Append new blocks
        payload="{\"children\": $blocks}"
        result=$(notion_api PATCH "/blocks/$page_id/children" "$payload")
        check_error "$result"

        echo "$result" | jq '{
            blocks_added: (.results | length),
            page_id: "'$page_id'"
        }'
        ;;

    append-body)
        # Append content to page (does not delete existing)
        # Usage: notion-query.sh append-body <page_id> <markdown>
        page_id="$2"
        markdown_input="$3"

        if [ -z "$page_id" ]; then
            echo "Usage: notion-query.sh append-body <page_id> <markdown>" >&2
            echo "       echo 'markdown' | notion-query.sh append-body <page_id> -" >&2
            exit 1
        fi

        # Read from stdin if - is passed
        if [ "$markdown_input" = "-" ]; then
            markdown_input=$(cat)
        fi

        if [ -z "$markdown_input" ]; then
            echo "Error: No markdown content provided" >&2
            exit 1
        fi

        # Convert markdown to Notion blocks JSON (uses markdown_to_rich_text from set-body)
        blocks="["
        first_block=true
        in_code=false
        code_lang=""
        code_content=""

        while IFS= read -r line || [ -n "$line" ]; do
            # Code block handling
            if [[ "$line" =~ ^\`\`\` ]]; then
                if [ "$in_code" = false ]; then
                    in_code=true
                    code_lang="${line:3}"
                    code_content=""
                else
                    in_code=false
                    [ "$first_block" = true ] || blocks+=","
                    first_block=false
                    code_content="${code_content//\\/\\\\}"
                    code_content="${code_content//\"/\\\"}"
                    [ -z "$code_lang" ] && code_lang="plain text"
                    blocks+="{\"object\":\"block\",\"type\":\"code\",\"code\":{\"language\":\"$code_lang\",\"rich_text\":[{\"type\":\"text\",\"text\":{\"content\":\"$code_content\"}}]}}"
                fi
                continue
            fi

            if [ "$in_code" = true ]; then
                [ -n "$code_content" ] && code_content+="\\n"
                escaped_line="${line//\\/\\\\}"
                escaped_line="${escaped_line//\"/\\\"}"
                code_content+="$escaped_line"
                continue
            fi

            # Skip empty lines
            [[ "$line" =~ ^[[:space:]]*$ ]] && continue

            [ "$first_block" = true ] || blocks+=","
            first_block=false

            # Heading 1
            if [[ "$line" =~ ^#\  ]]; then
                text="${line:2}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"heading_1\",\"heading_1\":{\"rich_text\":$rich_text}}"
            # Heading 2
            elif [[ "$line" =~ ^##\  ]]; then
                text="${line:3}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"heading_2\",\"heading_2\":{\"rich_text\":$rich_text}}"
            # Heading 3
            elif [[ "$line" =~ ^###\  ]]; then
                text="${line:4}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"heading_3\",\"heading_3\":{\"rich_text\":$rich_text}}"
            # Divider
            elif [[ "$line" =~ ^---$ ]] || [[ "$line" =~ ^\*\*\*$ ]] || [[ "$line" =~ ^___$ ]]; then
                blocks+="{\"object\":\"block\",\"type\":\"divider\",\"divider\":{}}"
            # Todo checked
            elif [[ "$line" =~ ^-\ \[[xX]\]\  ]]; then
                text="${line:6}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"to_do\",\"to_do\":{\"checked\":true,\"rich_text\":$rich_text}}"
            # Todo unchecked
            elif [[ "$line" =~ ^-\ \[\ \]\  ]]; then
                text="${line:6}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"to_do\",\"to_do\":{\"checked\":false,\"rich_text\":$rich_text}}"
            # Bullet list
            elif [[ "$line" =~ ^-\  ]] || [[ "$line" =~ ^\*\  ]]; then
                text="${line:2}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"bulleted_list_item\",\"bulleted_list_item\":{\"rich_text\":$rich_text}}"
            # Numbered list
            elif [[ "$line" =~ ^[0-9]+\.\  ]]; then
                text="${line#*. }"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"numbered_list_item\",\"numbered_list_item\":{\"rich_text\":$rich_text}}"
            # Quote
            elif [[ "$line" =~ ^\>\  ]]; then
                text="${line:2}"
                rich_text=$(markdown_to_rich_text "$text")
                blocks+="{\"object\":\"block\",\"type\":\"quote\",\"quote\":{\"rich_text\":$rich_text}}"
            # Regular paragraph
            else
                rich_text=$(markdown_to_rich_text "$line")
                blocks+="{\"object\":\"block\",\"type\":\"paragraph\",\"paragraph\":{\"rich_text\":$rich_text}}"
            fi
        done <<< "$markdown_input"

        blocks+="]"

        payload="{\"children\": $blocks}"
        result=$(notion_api PATCH "/blocks/$page_id/children" "$payload")
        check_error "$result"

        echo "$result" | jq '{
            blocks_added: (.results | length),
            page_id: "'$page_id'"
        }'
        ;;

    me)
        # Get current bot user info - compact
        result=$(notion_api GET "/users/me")
        check_error "$result"
        echo "$result" | jq '{
            id: .id,
            name: .name,
            type: .type
        }'
        ;;

    get-related)
        # Get all pages linked via a relation property
        # Usage: notion-query.sh get-related <page_id> [relation_property]
        # If relation_property is not specified, finds first relation property
        page_id="$2"
        relation_prop="$3"

        if [ -z "$page_id" ]; then
            echo "Usage: notion-query.sh get-related <page_id> [relation_property]" >&2
            exit 1
        fi

        # Get page with all properties
        page_result=$(notion_api GET "/pages/$page_id")
        check_error "$page_result"

        # Find relation property and extract IDs
        if [ -z "$relation_prop" ]; then
            # Auto-detect first relation property
            relation_data=$(echo "$page_result" | jq -r '
                .properties | to_entries[] | select(.value.type == "relation") |
                {name: .key, ids: [.value.relation[].id]}
            ' | head -1)
            relation_prop=$(echo "$relation_data" | jq -r '.name')
        fi

        # Get related page IDs
        related_ids=$(echo "$page_result" | jq -r --arg prop "$relation_prop" '
            .properties[$prop].relation // [] | .[].id
        ')

        if [ -z "$related_ids" ]; then
            echo "[]"
            exit 0
        fi

        # Fetch each related page and output compact info
        echo "["
        first=true
        while IFS= read -r rid; do
            [ -z "$rid" ] && continue
            [ "$first" = true ] || echo ","
            first=false

            related_page=$(notion_api GET "/pages/$rid")
            echo "$related_page" | jq '{
                id: .id,
                title: (.properties.Name.title[0].plain_text // .properties.title.title[0].plain_text // "Untitled"),
                status: (.properties.Status.status.name // null),
                url: .url
            }'
        done <<< "$related_ids"
        echo "]"
        ;;

    update-related)
        # Update all pages linked via a relation property
        # Usage: notion-query.sh update-related <page_id> <status> [relation_property]
        # Example: notion-query.sh update-related <page_id> "Done" Tasks
        page_id="$2"
        new_status="$3"
        relation_prop="${4:-Tasks}"

        if [ -z "$page_id" ] || [ -z "$new_status" ]; then
            echo "Usage: notion-query.sh update-related <page_id> <status> [relation_property]" >&2
            echo "Example: notion-query.sh update-related abc123 'Done' Tasks" >&2
            exit 1
        fi

        # Get related page IDs
        page_result=$(notion_api GET "/pages/$page_id")
        check_error "$page_result"

        related_ids=$(echo "$page_result" | jq -r --arg prop "$relation_prop" '
            .properties[$prop].relation // [] | .[].id
        ')

        if [ -z "$related_ids" ]; then
            echo "No related pages found in '$relation_prop' property"
            exit 0
        fi

        # Update each related page
        updated=0
        while IFS= read -r rid; do
            [ -z "$rid" ] && continue

            payload=$(jq -n --arg status "$new_status" '{properties: {Status: {status: {name: $status}}}}')
            result=$(notion_api PATCH "/pages/$rid" "$payload")

            if echo "$result" | jq -e '.object == "error"' > /dev/null 2>&1; then
                echo "Failed to update $rid: $(echo "$result" | jq -r '.message')" >&2
            else
                title=$(echo "$result" | jq -r '.properties.Name.title[0].plain_text // "Untitled"')
                echo "Updated: $title -> $new_status"
                updated=$((updated + 1))
            fi
        done <<< "$related_ids"

        echo "Updated $updated related pages"
        ;;

    *)
        echo "Notion API CLI (Token-Efficient)"
        echo ""
        echo "Usage: notion-query.sh <command> [args]"
        echo ""
        echo "READ Commands:"
        echo "  search [query]              Search pages/databases (compact)"
        echo "  find-db [query]             Find databases and show database IDs"
        echo "  query-db <db_id> [opts]     Query database (all properties)"
        echo "  backlog [roadmap]           Query configured backlog DB (shortcut)"
        echo "  get-page <page_id>          Get page properties (compact)"
        echo "  get-blocks <block_id>       Get block text content (JSON)"
        echo "  get-markdown <page_id>      Get page content as markdown"
        echo "  get-db <db_id>              Get database schema (property types)"
        echo "  get-related <page_id> [prop] Get all related pages (with status)"
        echo "  me                          Get bot info"
        echo ""
        echo "WRITE Commands:"
        echo "  create <db_id> <title>      Create page with any properties"
        echo "  create-backlog <title>      Create item in configured backlog DB"
        echo "  update <page_id>            Update page with any properties"
        echo "  update-related <id> <status> Update all related pages' status"
        echo "  archive <page_id>           Archive (trash) a page"
        echo "  set-body <page_id> <md>     Replace page content with markdown"
        echo "  append-body <page_id> <md>  Append markdown to page"
        echo ""
        echo "Generic Property Options (for create/update):"
        echo "  -p NAME=VALUE               Auto-detect type from schema"
        echo "  --title=X                   Set page title"
        echo "  --select-NAME=VALUE         Set select property"
        echo "  --status-NAME=VALUE         Set status property"
        echo "  --text-NAME=VALUE           Set rich_text property"
        echo "  --number-NAME=VALUE         Set number property"
        echo "  --checkbox-NAME=VALUE       Set checkbox (true/false)"
        echo "  --date-NAME=VALUE           Set date (YYYY-MM-DD)"
        echo "  --url-NAME=VALUE            Set URL property"
        echo "  --multi-select-NAME=VALUE   Add to multi_select"
        echo ""
        echo "Query Options:"
        echo "  --raw                       Return full API response"
        echo "  --filter-PROP=VALUE         Filter results by property"
        echo ""
        echo "Environment:"
        echo "  NOTION_API_KEY              Notion integration token (required)"
        echo "  NOTION_API_VERSION          Notion API version (default: 2022-06-28)"
        echo "  NOTION_BACKLOG_DB_ID        Optional DB ID for backlog/create-backlog shortcuts"
        echo ""
        echo "Examples:"
        echo "  # Find database ID by name"
        echo "  notion-query.sh find-db 'Roadmap'"
        echo ""
        echo "  # Get database schema first to see property names/types"
        echo "  notion-query.sh get-db <db_id>"
        echo ""
        echo "  # Create with auto-detected types"
        echo "  notion-query.sh create <db_id> 'New Task' -p 'Status=In progress' -p 'Priority=High'"
        echo ""
        echo "  # Create with explicit types"
        echo "  notion-query.sh create <db_id> 'Task' --status-Status='Not started' --select-Priority='Medium'"
        echo ""
        echo "  # Update any property"
        echo "  notion-query.sh update <page_id> -p 'Status=Done' --number-Score=95"
        echo ""
        echo "  # Filter query results"
        echo "  notion-query.sh query-db <db_id> --filter-Status='In progress'"
        echo ""
        echo "Markdown support (set-body/append-body):"
        echo "  # H1, ## H2, ### H3         Headings"
        echo "  - item, * item              Bullets"
        echo "  1. item                     Numbered"
        echo "  - [ ] todo, - [x] done      Checkboxes"
        echo "  > quote                     Blockquote"
        echo "  \`\`\`lang ... \`\`\`             Code blocks"
        echo "  ---                         Divider"
        ;;
esac
