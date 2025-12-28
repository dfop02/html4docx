"""
CSS Parser for HTML4DOCX

This module provides functionality to parse CSS from <style> tags and external CSS files.
It stores CSS rules by selector type (tag, class, id) and provides methods to retrieve
applicable styles for HTML elements.

This module is designed to be reusable for both <style> tags and external CSS files
loaded via <link> tags (future feature).
"""

import re
from typing import Dict, List, Tuple, Optional, Any


class CSSParser:
    """
    Parser for CSS rules from <style> tags or external CSS files.

    Stores CSS rules organized by selector type:
    - Tag selectors (e.g., 'p', 'h1', 'div')
    - Class selectors (e.g., '.my-class')
    - ID selectors (e.g., '#my-id')

    Supports basic CSS parsing including:
    - Simple selectors (tag, class, id)
    - Multiple selectors separated by commas
    - Style declarations with properties and values
    - !important flags
    """

    def __init__(self):
        """Initialize the CSS parser with empty rule storage."""
        # Store rules by selector type
        self.tag_rules: Dict[str, Dict[str, str]] = {}  # tag -> {property: value}
        self.class_rules: Dict[str, Dict[str, str]] = {}  # class -> {property: value}
        self.id_rules: Dict[str, Dict[str, str]] = {}  # id -> {property: value}

        # Store all rules with their specificity for proper cascade
        self._all_rules: List[Tuple[int, str, Dict[str, str]]] = []  # (specificity, selector, styles)

        # Track which elements are used in HTML (for selective CSS loading)
        self._used_tags: set = set()
        self._used_classes: set = set()
        self._used_ids: set = set()

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

        # Split by rules (look for selectors followed by { ... })
        # Pattern matches: selector { declarations }
        rule_pattern = re.compile(
            r'([^{]+)\{([^}]+)\}',
            re.MULTILINE | re.DOTALL
        )

        for match in rule_pattern.finditer(css_content):
            selectors_str = match.group(1).strip()
            declarations_str = match.group(2).strip()

            if not selectors_str or not declarations_str:
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

                # Store rule with specificity
                self._all_rules.append((specificity, selector, styles))

                # Store by selector type for quick lookup
                self._store_rule(selector, styles)

    def _remove_comments(self, css_content: str) -> str:
        """Remove CSS comments from content."""
        # Remove /* ... */ comments
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

        # Split by semicolon and parse each declaration
        for declaration in declarations.split(';'):
            declaration = declaration.strip()
            if not declaration or ':' not in declaration:
                continue

            # Split by first colon only (values may contain colons)
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

        # Count IDs
        id_count = len(re.findall(r'#[\w-]+', selector))
        specificity += id_count * 100

        # Count classes and attributes
        class_count = len(re.findall(r'\.[\w-]+', selector))
        attr_count = len(re.findall(r'\[[\w-]+\]', selector))
        specificity += (class_count + attr_count) * 10

        # Count tags
        tag_count = len(re.findall(r'^[\w-]+|(?<=\s)[\w-]+(?=\s|\.|#|\[|$)', selector))
        specificity += tag_count

        return specificity

    def _store_rule(self, selector: str, styles: Dict[str, str]) -> None:
        """
        Store CSS rule by selector type.

        Args:
            selector (str): CSS selector
            styles (Dict[str, str]): CSS properties and values
        """
        selector = selector.strip()

        # Handle ID selector (#id)
        if selector.startswith('#'):
            id_name = selector[1:].strip()
            if id_name:
                if id_name not in self.id_rules:
                    self.id_rules[id_name] = {}
                self.id_rules[id_name].update(styles)

        # Handle class selector (.class)
        elif selector.startswith('.'):
            class_name = selector[1:].strip()
            if class_name:
                if class_name not in self.class_rules:
                    self.class_rules[class_name] = {}
                self.class_rules[class_name].update(styles)

        # Handle tag selector (tag)
        else:
            # Remove any pseudo-classes or pseudo-elements
            tag_name = re.sub(r':[\w-]+', '', selector).strip()
            # Remove any combinators and keep only the tag
            tag_name = re.sub(r'[\s>+~].*', '', tag_name).strip()

            if tag_name:
                if tag_name not in self.tag_rules:
                    self.tag_rules[tag_name] = {}
                self.tag_rules[tag_name].update(styles)

    def get_styles_for_element(
        self,
        tag: str,
        attrs: Optional[Dict[str, str]] = None,
        inline_styles: Optional[Dict[str, str]] = None
    ) -> Dict[str, str]:
        """
        Get all applicable CSS styles for an HTML element.

        Combines styles from:
        1. Tag selectors
        2. Class selectors (from class attribute)
        3. ID selectors (from id attribute)
        4. Inline styles (highest priority)

        Styles are merged in order of specificity, with inline styles taking precedence.

        Args:
            tag (str): HTML tag name (e.g., 'p', 'div', 'span')
            attrs (Dict[str, str], optional): HTML attributes (e.g., {'class': 'my-class', 'id': 'my-id'})
            inline_styles (Dict[str, str], optional): Inline styles from style attribute

        Returns:
            Dict[str, str]: Combined CSS styles dictionary

        Example:
            parser = CSSParser()
            parser.parse_css("p { color: red; } .highlight { font-weight: bold; }")
            styles = parser.get_styles_for_element('p', {'class': 'highlight'})
            # Returns: {'color': 'red', 'font-weight': 'bold'}
        """
        combined_styles = {}

        if not attrs:
            attrs = {}

        # 1. Apply tag styles (lowest priority)
        if tag in self.tag_rules:
            combined_styles.update(self.tag_rules[tag])

        # 2. Apply class styles
        if 'class' in attrs:
            classes = attrs['class'].split()
            for class_name in classes:
                if class_name in self.class_rules:
                    combined_styles.update(self.class_rules[class_name])

        # 3. Apply ID styles
        if 'id' in attrs:
            element_id = attrs['id']
            if element_id in self.id_rules:
                combined_styles.update(self.id_rules[element_id])

        # 4. Apply inline styles (highest priority, except !important)
        if inline_styles:
            combined_styles.update(inline_styles)

        return combined_styles

    def get_styles_for_element_with_important(
        self,
        tag: str,
        attrs: Optional[Dict[str, str]] = None,
        inline_styles: Optional[Dict[str, str]] = None,
        inline_important: Optional[Dict[str, str]] = None
    ) -> Tuple[Dict[str, str], Dict[str, str]]:
        """
        Get styles separated by !important flag.

        Returns normal styles and important styles separately, following CSS cascade rules.

        Args:
            tag (str): HTML tag name
            attrs (Dict[str, str], optional): HTML attributes
            inline_styles (Dict[str, str], optional): Normal inline styles
            inline_important (Dict[str, str], optional): !important inline styles

        Returns:
            Tuple[Dict[str, str], Dict[str, str]]: (normal_styles, important_styles)
        """
        normal_styles = {}
        important_styles = {}

        if not attrs:
            attrs = {}

        # Sort all rules by specificity (ascending) to apply in correct order
        applicable_rules = []

        # Collect tag rules
        if tag in self.tag_rules:
            specificity = self._calculate_specificity(tag)
            applicable_rules.append((specificity, self.tag_rules[tag]))

        # Collect class rules
        if 'class' in attrs:
            classes = attrs['class'].split()
            for class_name in classes:
                if class_name in self.class_rules:
                    specificity = self._calculate_specificity(f'.{class_name}')
                    applicable_rules.append((specificity, self.class_rules[class_name]))

        # Collect ID rules
        if 'id' in attrs:
            element_id = attrs['id']
            if element_id in self.id_rules:
                specificity = self._calculate_specificity(f'#{element_id}')
                applicable_rules.append((specificity, self.id_rules[element_id]))

        # Apply rules in order of specificity (sort by first element - specificity)
        for specificity, styles in sorted(applicable_rules, key=lambda x: x[0]):
            for prop, value in styles.items():
                if '!important' in value.lower():
                    clean_value = value.replace('!important', '').strip()
                    important_styles[prop] = clean_value
                else:
                    normal_styles[prop] = value

        # Apply inline styles (normal)
        if inline_styles:
            normal_styles.update(inline_styles)

        # Apply inline !important styles (highest priority)
        if inline_important:
            important_styles.update(inline_important)

        return normal_styles, important_styles

    def clear(self) -> None:
        """Clear all stored CSS rules and used elements."""
        self.tag_rules.clear()
        self.class_rules.clear()
        self.id_rules.clear()
        self._all_rules.clear()
        self.clear_used_elements()

    def has_rules(self) -> bool:
        """Check if parser has any CSS rules stored."""
        return bool(self.tag_rules or self.class_rules or self.id_rules)

    def has_rules_for_element(self, tag: str, attrs: Optional[Dict[str, str]] = None) -> bool:
        """Check if parser has any CSS rules stored for an element."""
        if not self.has_rules() or not attrs:
            return False

        if tag in self.tag_rules:
            return True

        if 'class' in attrs:
            classes = attrs['class'].split()
            for class_name in classes:
                if class_name in self.class_rules:
                    return True

        if 'id' in attrs:
            element_id = attrs['id']
            if element_id in self.id_rules:
                return True

        return False

    def mark_element_used(self, tag: str, attrs: Optional[Dict[str, Any]] = None) -> None:
        """
        Mark an element as used in the HTML document.
        Used for selective CSS parsing to only load relevant rules.
        """
        if tag:
            self._used_tags.add(tag.lower())

        if not attrs:
            return

        # Handle class attribute (BeautifulSoup returns a list)
        classes = attrs.get('class', None)
        if classes:
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

        Uses precise matching to avoid false positives (e.g., "embed" matching "embed-responsive").

        Args:
            selector (str): CSS selector to check

        Returns:
            bool: True if selector matches any used element, False otherwise
        """
        if not selector:
            return False

        selector = selector.strip()

        # Check for ID selector (#id) - exact match
        id_matches = re.findall(r'#([\w-]+)', selector)
        if id_matches:
            for id_match in id_matches:
                if id_match in self._used_ids:
                    return True

        # Check for class selector (.class) - exact match or as part of class list
        class_matches = re.findall(r'\.([\w-]+)', selector)
        if class_matches:
            for class_match in class_matches:
                if class_match in self._used_classes:
                    return True

        # Check for tag selector (extract tag name)
        # Remove pseudo-classes, combinators, etc.
        tag_name = re.sub(r'[:#>+~\[].*', '', selector).strip()
        tag_name = re.sub(r'[#\.].*', '', tag_name).strip()
        tag_name = re.sub(r':.*', '', tag_name).strip()

        if tag_name and tag_name in self._used_tags:
            return True

        # Check for complex selectors like "div.container" or "p#header"
        # Split selector into parts and check each
        selector_parts = re.split(r'[#>+~\[\s,\.]', selector)
        for part in selector_parts:
            part = part.strip()
            if not part:
                continue

            # Exact match for tags
            if part in self._used_tags:
                return True

            # Exact match for classes (without the dot)
            if part in self._used_classes:
                return True

            # Exact match for IDs (without the hash)
            if part in self._used_ids:
                return True

        return False

    def clear_used_elements(self) -> None:
        """Clear the set of used elements (for selective parsing)."""
        self._used_tags.clear()
        self._used_classes.clear()
        self._used_ids.clear()

