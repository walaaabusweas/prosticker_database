import re


def parse_allowed(raw):
    if raw is None:
        return []
    s = str(raw).strip()
    if not s:
        return []
    # The schema uses slash-separated examples for most enum-like fields.
    return [x.strip() for x in re.split(r'\s*/\s*', s) if x.strip()]


def coerce_value(data_type, value):
    if value is None or value == '':
        return None
    dt = str(data_type or '').lower()
    if 'boolean' in dt:
        return bool(value) if isinstance(value, bool) else str(value).lower() in {'yes','true','1'}
    if 'number' in dt or 'integer' in dt:
        try:
            f = float(value)
            return int(f) if f.is_integer() else f
        except (ValueError, TypeError):
            raise ValueError('Must be a number.')
    return str(value).strip()


def validate_product(schema, data):
    errors = []
    final_type = data.get('product_type')
    subtype = data.get('subtype')
    if not final_type:
        errors.append('Product type is required.')
        return errors
    valid_types = {r.get('Final Type') for r in schema['final_types']}
    if final_type not in valid_types:
        errors.append('Invalid product type.')
        return errors
    valid_subtypes = {r.get('Subtype') for r in schema['subtypes'] if r.get('Final Type') == final_type}
    if subtype and subtype not in valid_subtypes:
        errors.append(f'Invalid subtype "{subtype}" for {final_type}.')
    if not str(data.get('product_name') or '').strip():
        errors.append('Product name is required.')

    for p in schema['properties']:
        if p.get('final_type') != final_type:
            continue
        if p.get('scope') == 'SUBTYPE_ONLY' and p.get('subtype') != subtype:
            continue
        if p.get('required') and not data.get(p['property']):
            errors.append(f"Required field: {p['property']}")
    return errors
