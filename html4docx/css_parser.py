"""
CSS Parser for HTML4DOCX

This module provides functionality to parse CSS from <style> tags and external CSS files.
It stores CSS rules by selector type (tag, class, id, compound) and provides methods to
retrieve applicable styles for HTML elements.

This module is designed to be reusable for both <style> tags and external CSS files
loaded via <link> tags.
"""

import re
from typing import Any, Dict, List, Optional, Tuple


class CSSParser:
    """
    Parser for CSS rules from <style> tags or external CSS files.

    Stores CSS rules organized by selector type:
    - Tag selectors (e.g., 'p', 'h1', 'div')
    - Class selectors (e.g., '.my-class')
    - ID selectors (e.g., '#my-id')
    - Compound selectors (e.g., 'p.my-class', 'div#header', '.a.b')

    Supports basic CSS parsing including:
    - Simple selectors (tag, class, id)
    - Compound selectors (tag+class, tag+id, multi-class)
    - Multiple selectors separated by commas
    - Style declarations with properties and values
    - !important flags
    - At-rule skipping (@media, @keyframes, @import, etc.)
    """

    def __init__(self):
        """Initialize the CSS parser with empty rule storage."""
        # Store rules by selector type (simple selectors)
        self.tag_rules: Dict[str, Dict[str, str]] = {}    # tag -> {property: value}
        self.class_rules: Dict[str, Dict[str, str]] = {}  # class -> {property: value}
        self.id_rules: Dict[str, Dict[str, str]] = {}     # id -> {property: value}

        # Store compound rules (tag+class, multi-class, tag+id, etc.)
        # Each entry: (specificity, tag_or_None, [classes], id_or_None, styles)
        self._compound_rules: List[Tuple[int, Optional[str], List[str], Optional[str], Dict[str, str]]] = []

        # Track which elements are used in HTML (for selective CSS loading)
        self._used_tags: set = set()
        self._used_classes: set = set()
        self._used_ids: set = set()

        # Store inline styles as temporary rules with very high specificity
        # Key: (tag, element_id) where element_id is unique per element instance
        # Value: (normal_styles, important_styles)
        self._inline_rules: Dict[Tuple[str, str], Tuple[Dict[str, str], Dict[str, str]]] = {}
        self._inline_rule_counter: int = 0

    def parse_css(self, css_content: str, selective: bool = False) -> None:
        """
        Parse CSS content and store rules by selector type.

        Args:
            css_content (str): CSS content from <style> tag or external file
            selective (bool): If True, only parse rules that match used elements
                             (tags, classes, IDs found in HTML). Default False.

        Example:
            parser = CSSParser()
            parser.parse_css("p { color: red; } .my-class { font-size: 12px; }")

            # Selective parsing (only load relevant rules)
            parser.mark_element_used('p', {'class': 'my-class'})
            parser.parse_css(large_css_file, selective=True)
        """
        if not css_content:
            return

        # Remove comments
        css_content = self._remove_comments(css_content)

        # Remove standalone @-rules without blocks first (e.g. @import, @charset).
        # Must be done BEFORE block at-rule removal to avoid [^{;]* crossing semicolons.
        css_content = re.sub(r'@[\w-]+[^;{]+;', '', css_content)
        # Remove @-rule blocks (e.g. @media, @keyframes, @supports).
        # [^{;]* stops at semicolons so we don't accidentally consume subsequent rules.
        css_content = re.sub(r'@[\w-]+[^{;]*\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', '', css_content, flags=re.DOTALL)

        # Split by rules (look for selectors followed by { ... })
        rule_pattern = re.compile(
            r'([^{]+)\{([^}]+)\}',
            re.MULTILINE | re.DOTALL
        )

        for match in rule_pattern.finditer(css_content):
            selectors_str = match.group(1).strip()
            declarations_str = match.group(2).strip()

            if not selectors_str or not declarations_str:
                continue

            # Skip leftover at-rule artifacts
            if selectors_str.lstrip().startswith('@'):
                continue

            # Parse declarations into dict
            styles = self._parse_declarations(declarations_str)
            if not styles:
                continue

            # Split selectors by comma (handle multiple selectors)
            selectors = [s.strip() for s in selectors_str.split(',')]

            for selector in selectors:
                if not selector:
                    continue

                # If selective parsing, check if this selector is relevant
                if selective and not self._is_selector_relevant(selector):
                    continue

                # Calculate specificity for cascade order
                specificity = self._calculate_specificity(selector)

                # Store by selector type for lookup
                self._store_rule(selector, styles, specificity)

    def _remove_comments(self, css_content: str) -> str:
        """Remove CSS comments from content."""
        return re.sub(r'/\*.*?\*/', '', css_content, flags=re.DOTALL)

    def _parse_declarations(self, declarations: str) -> Dict[str, str]:
        """
        Parse CSS declarations into a dictionary.

        Args:
            declarations (str): CSS declarations (e.g., "color: red; font-size: 12px")

        Returns:
            Dict[str, str]: Dictionary of property -> value
        """
        styles = {}

        for declaration in declarations.split(';'):
            declaration = declaration.strip()
            if not declaration or ':' not in declaration:
                continue

            parts = declaration.split(':', 1)
            if len(parts) != 2:
                continue

            property_name = parts[0].strip().lower()
            property_value = parts[1].strip()

            if property_name and property_value:
                styles[property_name] = property_value

        return styles

    def _calculate_specificity(self, selector: str) -> int:
        """
        Calculate CSS specificity for cascade order.

        Simple specificity calculation:
        - IDs: 100 points each
        - Classes/attributes: 10 points each
        - Tags: 1 point each

        Args:
            selector (str): CSS selector

        Returns:
            int: Specificity score (higher = more specific)
        """
        specificity = 0

        id_count = len(re.findall(r'#[\w-]+', selector))
        specificity += id_count * 100

        class_count = len(re.findall(r'\.[\w-]+', selector))
        attr_count = len(re.findall(r'\[[\w-]+\]', selector))
        specificity += (class_count + attr_count) * 10

        tag_count = len(re.findall(r'^[\w-]+|(?<=\s)[\w-]+(?=\s|\.|#|\[|$)', selector))
        specificity += tag_count

        return specificity

    def _parse_compound_selector(self, selector: str) -> Tuple[Optional[str], List[str], Optional[str], bool]:
        """
        Parse a CSS selector into its component parts.

        Returns (tag, classes, id, is_contextual) where:
        - tag: the tag name, or None if no tag restriction
        - classes: list of required class names
        - id: the required ID, or None
        - is_contextual: True if the selector uses combinators (space, >, +, ~)
          which require contextual matching this parser doesn't support

        Examples:
            'p'           -> ('p', [], None, False)
            '.cls'        -> (None, ['cls'], None, False)
            '#id'         -> (None, [], 'id', False)
            'p.cls'       -> ('p', ['cls'], None, False)
            'p#id'        -> ('p', [], 'id', False)
            'div.a.b'     -> ('div', ['a', 'b'], None, False)
            '.a.b'        -> (None, ['a', 'b'], None, False)
            'div > p'     -> (None, [], None, True)   # contextual
        """
        cleaned = selector.strip()

        # Remove pseudo-classes/elements (including function-style like :nth-child(2n+1))
        cleaned = re.sub(r'::?[\w-]+(?:\([^)]*\))?', '', cleaned).strip()

        # Check for combinators that indicate contextual selectors
        if re.search(r'[\s>+~]', cleaned):
            return None, [], None, True

        # Extract optional leading tag name
        tag_match = re.match(r'^([a-zA-Z][\w-]*)', cleaned)
        tag = tag_match.group(1) if tag_match else None

        # Extract all classes
        classes = re.findall(r'\.([\w-]+)', cleaned)

        # Extract first id (multiple IDs are invalid CSS; take first)
        id_matches = re.findall(r'#([\w-]+)', cleaned)
        element_id = id_matches[0] if id_matches else None

        return tag, classes, element_id, False

    def _store_rule(self, selector: str, styles: Dict[str, str], specificity: int = 0) -> None:
        """
        Store a CSS rule in the appropriate bucket based on selector type.

        Simple selectors (pure tag, pure class, pure id) go into fast-lookup dicts.
        Compound selectors (tag+class, tag+id, multi-class, etc.) go into
        _compound_rules for matching at retrieval time.
        Contextual selectors (descendant, child, sibling) fall back to storing
        only the first simple part as a best-effort approximation.

        Args:
            selector (str): CSS selector
            styles (Dict[str, str]): CSS properties and values
            specificity (int): Pre-computed specificity (0 means compute on demand)
        """
        selector = selector.strip()

        tag, classes, element_id, is_contextual = self._parse_compound_selector(selector)

        if is_contextual:
            # Contextual selector: best-effort — extract and store the first simple part
            # (e.g. 'div > p' stores under 'div'; context relationship is ignored)
            first_part = re.split(r'\s*[\s>+~]\s*', selector.strip())[0].strip()
            if first_part and first_part != selector:
                self._store_rule(first_part, styles, specificity)
            return

        is_pure_tag = bool(tag) and not classes and not element_id
        is_pure_class = not tag and len(classes) == 1 and not element_id
        is_pure_id = not tag and not classes and bool(element_id)

        if is_pure_tag:
            if tag not in self.tag_rules:
                self.tag_rules[tag] = {}
            self.tag_rules[tag].update(styles)

        elif is_pure_class:
            class_name = classes[0]
            if class_name not in self.class_rules:
                self.class_rules[class_name] = {}
            self.class_rules[class_name].update(styles)

        elif is_pure_id:
            if element_id not in self.id_rules:
                self.id_rules[element_id] = {}
            self.id_rules[element_id].update(styles)

        else:
            # Compound selector: store for matching at retrieval time
            spec = specificity if specificity else self._calculate_specificity(selector)
            self._compound_rules.append((spec, tag, classes, element_id, dict(styles)))

    def _match_compound_rules(
        self,
        tag: str,
        attrs: Dict[str, str]
    ) -> List[Tuple[int, Dict[str, str]]]:
        """
        Return a list of (specificity, styles) for compound rules that match this element.

        A compound rule matches if:
        - Its tag restriction (if any) equals the element's tag
        - All its required classes are present in the element's class list
        - Its id restriction (if any) equals the element's id
        """
        if not self._compound_rules:
            return []

        element_classes = set(attrs.get('class', '').split()) if attrs else set()
        element_id = attrs.get('id', '') if attrs else ''

        matches = []
        for (specificity, rule_tag, rule_classes, rule_id, rule_styles) in self._compound_rules:
            if rule_tag and rule_tag != tag:
                continue
            if rule_classes and not all(c in element_classes for c in rule_classes):
                continue
            if rule_id and rule_id != element_id:
                continue
            matches.append((specificity, rule_styles))
        return matches

    def get_styles_for_element(
        self,
        tag: str,
        attrs: Optional[Dict[str, str]] = None,
        inline_styles: Optional[Dict[str, str]] = None
    ) -> Dict[str, str]:
        """
        Get all applicable CSS styles for an HTML element.

        Combines styles from (in order of increasing priority):
        1. Tag selectors
        2. Compound selectors (tag+class, multi-class, etc.)
        3. Class selectors (from class attribute)
        4. ID selectors (from id attribute)
        5. Inline styles (highest priority)

        Args:
            tag (str): HTML tag name (e.g., 'p', 'div', 'span')
            attrs (Dict[str, str], optional): HTML attributes
            inline_styles (Dict[str, str], optional): Inline styles from style attribute

        Returns:
            Dict[str, str]: Combined CSS styles dictionary
        """
        combined_styles = {}

        if not attrs:
            attrs = {}

        # 1. Apply tag styles (lowest priority)
        if tag in self.tag_rules:
            combined_styles.update(self.tag_rules[tag])

        # 2. Apply compound rules (specificity-ordered, then class/id on top)
        for _, rule_styles in sorted(self._match_compound_rules(tag, attrs), key=lambda x: x[0]):
            combined_styles.update(rule_styles)

        # 3. Apply class styles
        if 'class' in attrs:
            for class_name in attrs['class'].split():
                if class_name in self.class_rules:
                    combined_styles.update(self.class_rules[class_name])

        # 4. Apply ID styles
        if 'id' in attrs:
            element_id = attrs['id']
            if element_id in self.id_rules:
                combined_styles.update(self.id_rules[element_id])

        # 5. Apply inline styles (highest priority)
        if inline_styles:
            combined_styles.update(inline_styles)

        return combined_styles

    def add_inline_styles(self, tag: str, attrs: Optional[Dict[str, str]] = None,
                          inline_normal: Optional[Dict[str, str]] = None,
                          inline_important: Optional[Dict[str, str]] = None) -> str:
        """
        Add inline styles as temporary rules with very high specificity.

        Inline styles have higher specificity than ID selectors (1000+ points).
        Returns a unique element_id that can be used to remove these rules later.

        Args:
            tag (str): HTML tag name
            attrs (Dict[str, str], optional): HTML attributes
            inline_normal (Dict[str, str], optional): Normal inline styles
            inline_important (Dict[str, str], optional): !important inline styles

        Returns:
            str: Unique element_id for this inline rule (can be used to remove it)
        """
        if not inline_normal and not inline_important:
            return None

        if not attrs:
            attrs = {}

        self._inline_rule_counter += 1
        element_id = f"__inline_{self._inline_rule_counter}"

        self._inline_rules[(tag, element_id)] = (
            inline_normal or {},
            inline_important or {}
        )

        if 'data-inline-id' not in attrs:
            attrs['data-inline-id'] = element_id

        return element_id

    def remove_inline_styles(self, element_id: str) -> None:
        """Remove inline styles for a specific element."""
        if not element_id:
            return
        keys_to_remove = [k for k in self._inline_rules if k[1] == element_id]
        for key in keys_to_remove:
            self._inline_rules.pop(key, None)

    def get_styles_for_element_with_important(
        self,
        tag: str,
        attrs: Optional[Dict[str, str]] = None
    ) -> Tuple[Dict[str, str], Dict[str, str]]:
        """
        Get styles separated by !important flag.

        This is the single source of truth for all styles:
        - CSS rules from files and <style> tags
        - Inline styles (stored as temporary rules with high specificity)

        Applies the full CSS cascade:
        1. Tag rules (lowest specificity)
        2. Compound rules (tag+class, multi-class, etc.)
        3. Class rules (medium specificity)
        4. ID rules (high specificity)
        5. Inline styles (highest specificity — always wins)

        Returns normal styles and important styles separately.

        Args:
            tag (str): HTML tag name
            attrs (Dict[str, str], optional): HTML attributes (may contain 'data-inline-id')

        Returns:
            Tuple[Dict[str, str], Dict[str, str]]: (normal_styles, important_styles)
        """
        normal_styles = {}
        important_styles = {}

        if not attrs:
            attrs = {}

        # Build list of (specificity, styles_dict) for all applicable rules
        applicable_rules = []

        # 1. Tag rules (lowest specificity)
        if tag in self.tag_rules:
            specificity = self._calculate_specificity(tag)
            applicable_rules.append((specificity, self.tag_rules[tag]))

        # 2. Compound rules
        applicable_rules.extend(self._match_compound_rules(tag, attrs))

        # 3. Class rules
        if 'class' in attrs:
            for class_name in attrs['class'].split():
                if class_name in self.class_rules:
                    specificity = self._calculate_specificity(f'.{class_name}')
                    applicable_rules.append((specificity, self.class_rules[class_name]))

        # 4. ID rules (high specificity)
        if 'id' in attrs:
            element_id = attrs['id']
            if element_id in self.id_rules:
                specificity = self._calculate_specificity(f'#{element_id}')
                applicable_rules.append((specificity, self.id_rules[element_id]))

        # Apply rules in ascending specificity order (higher specificity wins)
        for _specificity, styles in sorted(applicable_rules, key=lambda x: x[0]):
            for prop, value in styles.items():
                if '!important' in value.lower():
                    clean_value = re.sub(r'!important', '', value, flags=re.IGNORECASE).strip()
                    important_styles[prop] = clean_value
                else:
                    normal_styles[prop] = value

        # 5. Apply inline styles (highest specificity — overrides everything)
        inline_id = attrs.get('data-inline-id')
        if inline_id:
            for (rule_tag, rule_id), (inline_normal, inline_important) in self._inline_rules.items():
                if rule_id == inline_id and rule_tag == tag:
                    if inline_normal:
                        normal_styles.update(inline_normal)
                    if inline_important:
                        important_styles.update(inline_important)
                    break

        return normal_styles, important_styles

    def clear(self) -> None:
        """Clear all stored CSS rules and used elements."""
        self.tag_rules.clear()
        self.class_rules.clear()
        self.id_rules.clear()
        self._compound_rules.clear()
        self._inline_rules.clear()
        self._inline_rule_counter = 0
        self.clear_used_elements()

    def has_rules(self) -> bool:
        """Check if parser has any CSS rules stored."""
        return bool(self.tag_rules or self.class_rules or self.id_rules or self._compound_rules)

    def has_rules_for_element(self, tag: str, attrs: Optional[Dict[str, str]] = None) -> bool:
        """Check if parser has any CSS rules stored for an element."""
        if not self.has_rules():
            return False

        # Tag rules exist for this element?
        if tag in self.tag_rules:
            return True

        # No attrs to check further
        if not attrs:
            return False

        # Class rules?
        if 'class' in attrs:
            for class_name in attrs['class'].split():
                if class_name in self.class_rules:
                    return True

        # ID rules?
        if 'id' in attrs and attrs['id'] in self.id_rules:
            return True

        # Compound rules?
        return bool(self._match_compound_rules(tag, attrs))

    def mark_element_used(self, tag: str, attrs: Optional[Dict[str, Any]] = None) -> None:
        """
        Mark an element as used in the HTML document.
        Used for selective CSS parsing to only load relevant rules.
        """
        if tag:
            self._used_tags.add(tag.lower())

        if not attrs:
            return

        # Handle class attribute — BeautifulSoup returns a list, regex fallback returns a string
        classes = attrs.get('class', None)
        if classes:
            if isinstance(classes, str):
                classes = classes.split()
            for class_name in classes:
                if class_name:
                    self._used_classes.add(class_name)

        # Handle id attribute (string)
        element_id = attrs.get('id', None)
        if element_id:
            self._used_ids.add(element_id)

    def _is_selector_relevant(self, selector: str) -> bool:
        """
        Check if a CSS selector is relevant based on used elements.

        Args:
            selector (str): CSS selector to check

        Returns:
            bool: True if selector matches any used element, False otherwise
        """
        if not selector:
            return False

        selector = selector.strip()

        # Check for ID selector (#id) — exact match
        id_matches = re.findall(r'#([\w-]+)', selector)
        if id_matches:
            for id_match in id_matches:
                if id_match in self._used_ids:
                    return True

        # Check for class selector (.class) — exact match
        class_matches = re.findall(r'\.([\w-]+)', selector)
        if class_matches:
            for class_match in class_matches:
                if class_match in self._used_classes:
                    return True

        # Check for tag selector — strip pseudo-classes, ids, classes, combinators
        tag_name = re.sub(r'[:#>+~\[].*', '', selector).strip()
        tag_name = re.sub(r'[#\.].*', '', tag_name).strip()
        tag_name = re.sub(r':.*', '', tag_name).strip()

        if tag_name and tag_name in self._used_tags:
            return True

        # Check parts of complex selectors
        selector_parts = re.split(r'[#>+~\[\s,\.]', selector)
        for part in selector_parts:
            part = part.strip()
            if not part:
                continue
            if part in self._used_tags:
                return True
            if part in self._used_classes:
                return True
            if part in self._used_ids:
                return True

        return False

    def clear_used_elements(self) -> None:
        """Clear the set of used elements (for selective parsing)."""
        self._used_tags.clear()
        self._used_classes.clear()
        self._used_ids.clear()
