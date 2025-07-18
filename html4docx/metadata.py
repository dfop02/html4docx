import json
from datetime import datetime
from docx import Document

class Metadata():
    """Handle docx document metadata"""
    def __init__(self, document: Document):
        self.document = document

    def __str__(self):
        self.get_metadata(print_result=True)

    def set_metadata(self, **kwargs) -> None:
        """
        Set metadata on docx document.

        Args:
            kwargs (dict): A dictionary of core properties from document.
                          Example: {"author": "Myself", "revision": "2", "keywords": {"custom_field":"123"}}
        sources:
            Core Properties: https://python-docx.readthedocs.io/en/latest/dev/analysis/features/coreprops.html
        """
        core_props = self.document.core_properties

        for key, value in kwargs.items():
            if hasattr(core_props, key):
                if key == 'revision':
                    try:
                        value = int(value)
                    except ValueError:
                        print(f'Invalid revision number "{value}". Must be an integer. Skipping...')
                        continue
                elif key in ['last_printed', 'modified', 'created']:
                    try:
                        value = datetime.fromisoformat(value)
                    except ValueError:
                        print(f'Invalid datetime string on property "{key}", must be in ISO format. Skipping...')
                        continue

                setattr(core_props, key, value)
            else:
                print(f'Property "{key}" not found in core properties. Skipping...')

    def get_metadata(self, print_result: bool = False):
        """
        Get metadata from docx document.

        Args:
            print_result (bool): Print the result instead of returning a dictionary.

        Returns:
            dict: Document core properties.

        Sources:
            Core Properties: https://python-docx.readthedocs.io/en/latest/dev/analysis/features/coreprops.html
        """
        core_props = self.document.core_properties

        props = {
            attr: getattr(core_props, attr)
            for attr in dir(core_props)
            if not callable(getattr(core_props, attr)) and not attr.startswith("_")
        }

        if print_result:
            print(json.dumps(props, sort_keys=True, indent=4, default=str))
        else:
            return props
