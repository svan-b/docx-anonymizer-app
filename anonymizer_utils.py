#!/usr/bin/env python3
"""
Shared Anonymization Utilities
Used by Word, PowerPoint, and Excel processors to avoid code duplication
"""


def anonymize_text(text, alias_map, sorted_keys, compiled_patterns=None, track_details=False):
    """
    Apply anonymization replacements with case matching using SINGLE-PASS regex (v2.1).

    Universal function used by all document processors (Word, PowerPoint, Excel).

    Args:
        text: Text to anonymize
        alias_map: Dictionary of original → replacement mappings
        sorted_keys: Sorted list of alias_map keys
        compiled_patterns: Pre-compiled regex patterns dict (if None, will compile on-the-fly)
        track_details: If True, track which originals were replaced

    Returns:
        If track_details=False: (text, replacements)
        If track_details=True: (text, replacements, details_dict)
    """
    if not text or not isinstance(text, str):
        if track_details:
            return text, 0, {}
        return text, 0

    replacements = 0

    # BACKWARD COMPATIBILITY: If patterns not pre-compiled, compile them now
    # Note: This requires access to precompile_patterns from process_adobe_word_files
    # For now, we'll just handle the case where compiled_patterns exists
    if compiled_patterns is None:
        # Cannot compile patterns here without importing precompile_patterns
        # This should only happen in legacy code paths
        raise ValueError("compiled_patterns cannot be None - please pass pre-compiled patterns")

    # Extract combined pattern and lookup map
    combined_pattern = compiled_patterns.get('combined')
    lookup = compiled_patterns.get('lookup')

    # BACKWARD COMPATIBILITY: Handle old compiled_patterns format
    if combined_pattern is None or lookup is None:
        # Old format - use legacy multi-pass algorithm
        result = anonymize_text_legacy(text, alias_map, sorted_keys, compiled_patterns)
        if track_details:
            return result[0], result[1], {}
        return result

    # Track which originals were replaced (v2.1 feature)
    details = {} if track_details else None

    # SINGLE-PASS REPLACEMENT (v2.1 performance optimization)
    def replace_match(match):
        nonlocal replacements
        matched_text = match.group(0)

        # Look up the replacement using lowercase match
        matched_lower = matched_text.lower()
        if matched_lower not in lookup:
            return matched_text  # Safe fallback

        original, replacement = lookup[matched_lower]

        # Track this replacement (v2.1)
        if track_details:
            details[original] = details.get(original, 0) + 1

        # Preserve case pattern
        if matched_text.isupper():
            replacements += 1
            return replacement.upper()
        elif matched_text.islower():
            replacements += 1
            return replacement.lower()
        elif matched_text[0].isupper():
            replacements += 1
            return replacement.capitalize()
        else:
            replacements += 1
            return replacement

    # Single regex pass replaces ALL patterns at once
    text = combined_pattern.sub(replace_match, text)

    if track_details:
        return text, replacements, details
    return text, replacements


def anonymize_text_legacy(text, alias_map, sorted_keys, compiled_patterns):
    """
    Legacy multi-pass anonymization (kept for backward compatibility).

    Used when compiled_patterns doesn't have the new 'combined' format.
    """
    if not text or not isinstance(text, str):
        return text, 0

    replacements = 0

    for original in sorted_keys:
        replacement = alias_map[original]

        def replace_with_case(match):
            nonlocal replacements
            matched_text = match.group(0)

            # Preserve case pattern
            if matched_text.isupper():
                replacements += 1
                return replacement.upper()
            elif matched_text.islower():
                replacements += 1
                return replacement.lower()
            elif matched_text[0].isupper():
                replacements += 1
                return replacement.capitalize()
            else:
                replacements += 1
                return replacement

        pattern = compiled_patterns[original]
        text = pattern.sub(replace_with_case, text)

    return text, replacements


def merge_details(details1, details2):
    """
    Merge two replacement details dictionaries (v2.1 helper).

    Args:
        details1: First details dict {original: count, ...}
        details2: Second details dict to merge in

    Returns:
        Merged details dict
    """
    if details1 is None:
        return details2 if details2 else {}
    if details2 is None:
        return details1

    merged = details1.copy()
    for original, count in details2.items():
        merged[original] = merged.get(original, 0) + count
    return merged
