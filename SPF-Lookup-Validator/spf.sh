#!/bin/bash

# SPF Record Analyzer & RFC 7208 Validator
# Usage: ./spf_parser.sh

max_depth=10
visited_domains=""
temp_file="/tmp/spf_lookups_$$"
warnings_file="/tmp/spf_warnings_$$"
errors_file="/tmp/spf_errors_$$"

# Initialize temp files
echo "0" > "$temp_file"
echo "" > "$warnings_file"
echo "" > "$errors_file"

# Function to add warning with chain
add_warning() {
    local message=$1
    local chain=$2
    if [ -n "$chain" ]; then
        echo "$message (chain: $chain)" >> "$warnings_file"
    else
        echo "$message" >> "$warnings_file"
    fi
}

# Function to add error with chain
add_error() {
    local message=$1
    local chain=$2
    if [ -n "$chain" ]; then
        echo "$message (chain: $chain)" >> "$errors_file"
    else
        echo "$message" >> "$errors_file"
    fi
}

# Function to increment lookup count
increment_lookups() {
    local count=$(cat "$temp_file")
    echo $((count + 1)) > "$temp_file"
}

# Function to get lookup count
get_lookups() {
    cat "$temp_file"
}

# Function to check if domain was visited
is_visited() {
    local domain=$1
    if echo "$visited_domains" | grep -q "^${domain}$"; then
        return 0
    else
        return 1
    fi
}

# Function to mark domain as visited
mark_visited() {
    local domain=$1
    visited_domains="${visited_domains}${domain}"$'\n'
}

# Function to clean SPF record - remove extra spaces
clean_spf_record() {
    local spf=$1
    # Remove multiple consecutive spaces and replace with single space
    echo "$spf" | tr -s ' '
}

# Function to validate SPF record format
validate_spf_format() {
    local spf_record=$1
    local domain=$2
    local chain=$3
    
    # Check if starts with v=spf1
    if ! echo "$spf_record" | grep -qE "^v=spf1( |$)"; then
        add_error "SPF record must start with 'v=spf1'" "$chain"
    fi
    
    # Check total length (RFC 7208: DNS TXT record should not exceed 512 bytes)
    local length=${#spf_record}
    if [ $length -gt 512 ]; then
        add_error "SPF record exceeds 512 bytes ($length bytes) - may cause DNS issues" "$chain"
    elif [ $length -gt 450 ]; then
        add_warning "SPF record is getting long ($length bytes) - consider staying under 450 bytes" "$chain"
    fi
    
    # Check for deprecated 'ptr' mechanism
    if echo "$spf_record" | grep -qE '(^| )[+~?-]?ptr(:|$| )'; then
        add_warning "DEPRECATED: 'ptr' mechanism found - RFC 7208 recommends not using ptr" "$chain"
    fi
    
    # Check for redirect and all together
    if echo "$spf_record" | grep -qE 'redirect=' && echo "$spf_record" | grep -qE '[+~?-]all'; then
        add_warning "Both 'redirect=' and 'all' found - 'all' will be ignored when redirect is processed" "$chain"
    fi
    
    # Validate IP addresses format - basic validation
    local ip4_addrs=$(echo "$spf_record" | grep -oE 'ip4:[^ ]+' | sed 's/ip4://')
    if [ -n "$ip4_addrs" ]; then
        while IFS= read -r ip; do
            if [ -n "$ip" ]; then
                # Basic IPv4 validation - allow CIDR notation
                if ! echo "$ip" | grep -qE '^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}(/[0-9]{1,2})?$'; then
                    add_error "Invalid IPv4 address format: $ip" "$chain"
                fi
            fi
        done <<< "$ip4_addrs"
    fi
}

# Function to fetch SPF record for a domain
fetch_spf_record() {
    local domain=$1
    local chain=$2
    local spf_records=$(dig +short "$domain" TXT 2>/dev/null | grep "v=spf1")
    local count=$(echo "$spf_records" | grep -c "v=spf1")
    
    # Check for multiple SPF records
    if [ $count -gt 1 ]; then
        add_error "Multiple SPF records found - RFC 7208 allows only ONE SPF record per domain" "$chain"
    fi
    
    # Clean the SPF record (remove quotes and normalize spaces)
    local cleaned=$(echo "$spf_records" | head -1 | tr -d '"' | tr -s ' ')
    echo "$cleaned"
}

# Function to parse and analyze SPF record recursively
analyze_spf() {
    local input=$1
    local depth=${2:-0}
    local parent_chain=${3:-""}
    local indent=""
    local spf_record=""
    local domain=""
    local is_spf_string=0
    local current_chain=""
    
    # Create indentation for nested output
    for ((i=0; i<depth; i++)); do
        indent="  $indent"
    done
    
    # Check if we've exceeded max depth
    if [ $depth -ge $max_depth ]; then
        echo "${indent}WARNING: Max recursion depth reached"
        add_warning "Maximum recursion depth of $max_depth reached - possible circular reference" "$parent_chain"
        return
    fi
    
    # Check if input is an SPF record or a domain
    if echo "$input" | grep -q "^v=spf1"; then
        # It's an SPF record - clean it first
        spf_record=$(clean_spf_record "$input")
        domain="(main SPF)"
        is_spf_string=1
        current_chain="$domain"
        
        # Validate format
        validate_spf_format "$spf_record" "$domain" "$current_chain"
    else
        # It's a domain
        domain="$input"
        
        # Build chain
        if [ -z "$parent_chain" ]; then
            current_chain="$domain"
        else
            current_chain="$parent_chain > $domain"
        fi
        
        # Check if already visited (for loop prevention)
        if is_visited "$domain"; then
            echo ""
            echo "${indent}Domain: $domain (already visited - counting lookup but not re-parsing)"
            increment_lookups
            return
        fi
        
        mark_visited "$domain"
        
        # Fetch SPF record (already cleaned in fetch function)
        spf_record=$(fetch_spf_record "$domain" "$current_chain")
        
        if [ -z "$spf_record" ]; then
            echo "${indent}Domain: $domain - NO SPF RECORD FOUND"
            add_warning "No SPF record found for domain: $domain" "$current_chain"
            increment_lookups
            return
        fi
        
        # Validate format
        validate_spf_format "$spf_record" "$domain" "$current_chain"
        
        # Increment lookup count for this domain lookup
        increment_lookups
    fi
    
    echo ""
    echo "${indent}Domain: $domain"
    echo "${indent}SPF Record: $spf_record"
    
    # Parse includes
    local includes=$(echo "$spf_record" | grep -oE 'include:[^ ]+' | sed 's/include://')
    if [ -n "$includes" ]; then
        echo "${indent}Includes:"
        echo "$includes" | while IFS= read -r inc_domain; do
            [ -n "$inc_domain" ] && echo "${indent}  - $inc_domain"
        done
    fi
    
    # Parse IPv4 addresses
    local ip4_addrs=$(echo "$spf_record" | grep -oE 'ip4:[^ ]+' | sed 's/ip4://')
    if [ -n "$ip4_addrs" ]; then
        echo "${indent}IPv4 Addresses:"
        echo "$ip4_addrs" | while IFS= read -r ip; do
            [ -n "$ip" ] && echo "${indent}  - $ip"
        done
    fi
    
    # Parse IPv6 addresses
    local ip6_addrs=$(echo "$spf_record" | grep -oE 'ip6:[^ ]+' | sed 's/ip6://')
    if [ -n "$ip6_addrs" ]; then
        echo "${indent}IPv6 Addresses:"
        echo "$ip6_addrs" | while IFS= read -r ip; do
            [ -n "$ip" ] && echo "${indent}  - $ip"
        done
    fi
    
    # Check for 'a' mechanism (counts as DNS lookup)
    local a_count=$(echo "$spf_record" | grep -oE '(^| )[+~?-]?a(:|$| )' | wc -l | tr -d ' ')
    if [ "$a_count" -gt 0 ]; then
        echo "${indent}'a' mechanism: allows domain's A record IPs (counts as $a_count lookup(s))"
        for ((i=0; i<a_count; i++)); do
            increment_lookups
        done
    fi
    
    # Check for 'mx' mechanism (counts as DNS lookup)
    local mx_count=$(echo "$spf_record" | grep -oE '(^| )[+~?-]?mx(:|$| )' | wc -l | tr -d ' ')
    if [ "$mx_count" -gt 0 ]; then
        echo "${indent}'mx' mechanism: allows domain's MX servers (counts as $mx_count lookup(s))"
        for ((i=0; i<mx_count; i++)); do
            increment_lookups
        done
    fi
    
    # Check for 'ptr' mechanism (deprecated)
    local ptr_count=$(echo "$spf_record" | grep -oE '(^| )[+~?-]?ptr(:|$| )' | wc -l | tr -d ' ')
    if [ "$ptr_count" -gt 0 ]; then
        echo "${indent}'ptr' mechanism (DEPRECATED): counts as $ptr_count lookup(s)"
        for ((i=0; i<ptr_count; i++)); do
            increment_lookups
        done
    fi
    
    # Check for 'exists' mechanism (counts as DNS lookup)
    local exists_count=$(echo "$spf_record" | grep -oE 'exists:[^ ]+' | wc -l | tr -d ' ')
    if [ "$exists_count" -gt 0 ]; then
        echo "${indent}'exists' mechanism found (counts as $exists_count lookup(s))"
        for ((i=0; i<exists_count; i++)); do
            increment_lookups
        done
    fi
    
    # Check for 'redirect' modifier (counts as DNS lookup)
    local redirect=$(echo "$spf_record" | grep -oE 'redirect=[^ ]+' | sed 's/redirect=//')
    if [ -n "$redirect" ]; then
        echo "${indent}'redirect' to: $redirect (counts as 1 lookup)"
        increment_lookups
    fi
    
    # Parse 'all' qualifier - dash at the end of character class
    local all_match=$(echo "$spf_record" | grep -oE '[+~?-]all')
    if [ -n "$all_match" ]; then
        local all_qualifier="${all_match:0:1}"
        echo "${indent}'all' mechanism qualifier: $all_qualifier"
        case $all_qualifier in
            "+") echo "${indent}  (Pass: allow all - NOT RECOMMENDED)" 
                 add_warning "'+all' allows ANY server to send email - this is insecure" "$current_chain";;
            "-") echo "${indent}  (Fail: reject all)" ;;
            "~") echo "${indent}  (SoftFail: accept but mark)" ;;
            "?") echo "${indent}  (Neutral: no policy)" 
                 add_warning "'?all' provides no protection - consider using '~all' or '-all'" "$current_chain";;
        esac
    else
        # Only warn if no redirect modifier exists
        if [ -z "$redirect" ]; then
            add_warning "No 'all' mechanism found - SPF record should end with an 'all' qualifier" "$current_chain"
        fi
    fi
    
    # Recursively analyze includes
    if [ -n "$includes" ]; then
        while IFS= read -r inc_domain; do
            if [ -n "$inc_domain" ]; then
                analyze_spf "$inc_domain" $((depth + 1)) "$current_chain"
            fi
        done <<< "$includes"
    fi
    
    # Handle redirect recursively
    if [ -n "$redirect" ]; then
        analyze_spf "$redirect" $((depth + 1)) "$current_chain"
    fi
}

# Main script
echo "=========================================="
echo "SPF Record Analyzer & RFC 7208 Validator"
echo "=========================================="
echo ""
read -p "Enter domain or full SPF record to check: " input

if [ -z "$input" ]; then
    echo "Error: No input provided"
    exit 1
fi

# Reset counters
echo "0" > "$temp_file"
echo "" > "$warnings_file"
echo "" > "$errors_file"
visited_domains=""

# Check if input is an SPF record string
if echo "$input" | grep -q "^v=spf1"; then
    # Add 1 for the initial domain lookup that would happen in real SPF validation
    increment_lookups
    echo "Note: Adding +1 lookup for initial domain TXT record query"
fi

# Analyze the SPF record
analyze_spf "$input"

# Get final lookup count
total_lookups=$(get_lookups)

# Display summary
echo ""
echo "=========================================="
echo "SUMMARY"
echo "=========================================="
echo "Total DNS lookups estimated: $total_lookups"
echo ""

# Check DNS lookup limit
if [ $total_lookups -gt 10 ]; then
    echo "❌ ERROR: DNS lookup count EXCEEDS the SPF limit of 10!"
    echo "   This WILL cause SPF PermError and email delivery failures."
    echo "   Lookups over limit: $((total_lookups - 10))"
elif [ $total_lookups -eq 10 ]; then
    echo "⚠️  CAUTION: DNS lookup count is exactly at the limit of 10."
    echo "   Consider reducing lookups to provide a safety margin."
else
    echo "✅ DNS lookup count is within the recommended limit of 10."
    echo "   Remaining lookups available: $((10 - total_lookups))"
fi
echo ""

# Display errors
if [ -s "$errors_file" ]; then
    echo "=========================================="
    echo "ERRORS FOUND"
    echo "=========================================="
    cat "$errors_file"
    echo ""
fi

# Display warnings
if [ -s "$warnings_file" ]; then
    echo "=========================================="
    echo "WARNINGS"
    echo "=========================================="
    cat "$warnings_file"
    echo ""
fi

# Final verdict
if [ ! -s "$errors_file" ] && [ ! -s "$warnings_file" ] && [ $total_lookups -le 10 ]; then
    echo "=========================================="
    echo "✅ SPF RECORD APPEARS VALID"
    echo "=========================================="
    echo "No errors or warnings found. SPF record complies with RFC 7208."
    echo ""
elif [ -s "$errors_file" ]; then
    echo "=========================================="
    echo "❌ SPF RECORD HAS ERRORS"
    echo "=========================================="
    echo "Please fix the errors above before deploying this SPF record."
    echo ""
else
    echo "=========================================="
    echo "⚠️  SPF RECORD HAS WARNINGS"
    echo "=========================================="
    echo "Consider addressing the warnings above to improve SPF compliance."
    echo ""
fi

# Cleanup
rm -f "$temp_file" "$warnings_file" "$errors_file"
