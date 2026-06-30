import re


class PlaceholderService:
    """Render {{Column_Name}} placeholders from dynamically discovered row data."""

    _pattern = re.compile(r"{{\s*([^{}]+?)\s*}}", re.IGNORECASE)

    @staticmethod
    def create_context(headers, values):
        context = {}
        for header, value in zip(headers, values):
            if header is None:
                continue
            key = str(header).strip().casefold()
            context[key] = "" if value is None else str(value)
        return context

    @classmethod
    def render(cls, template, context):
        text = "" if template is None else str(template)
        normalized = {str(key).strip().casefold(): "" if value is None else str(value) for key, value in context.items()}

        def replace(match):
            return normalized.get(match.group(1).strip().casefold(), "")

        return cls._pattern.sub(replace, text)

    @classmethod
    def render_row(cls, template, headers, values):
        return cls.render(template, cls.create_context(headers, values))
