"""
Per-property configuration, same pattern as the ga-automation pipeline's
data/{property_code}/config.yaml: cost centers are property-specific
(Riverside Labs has one, "west20"; Lexington Labs has five, "lexlab-1"
through "lexlab-5", tagged B1-B5 in Kardin's filenames) and shouldn't be
hardcoded or re-typed by hand every run.

Not required - app.py works exactly as before with no property selected.
Loading a config just auto-fills batch mode's per-building names and lets
kardin_parser cross-check that a building's files are actually scoped to
the cost center they're supposed to be (see kardin_parser.
check_multiple_cost_centers's expected_cost_center param).
"""
import os

import yaml

DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')


def list_properties():
    """[{slug, name}] for every data/{slug}/config.yaml found (TEMPLATE excluded)."""
    if not os.path.isdir(DATA_DIR):
        return []
    out = []
    for slug in sorted(os.listdir(DATA_DIR)):
        if slug == 'TEMPLATE':
            continue
        path = os.path.join(DATA_DIR, slug, 'config.yaml')
        if os.path.isfile(path):
            cfg = load_property_config(slug)
            out.append({'slug': slug, 'name': cfg.get('property_name', slug)})
    return out


def load_property_config(slug):
    path = os.path.join(DATA_DIR, slug, 'config.yaml')
    with open(path, 'r', encoding='utf-8') as f:
        cfg = yaml.safe_load(f) or {}
    cfg.setdefault('cost_centers', [])
    return cfg


def cost_center_for_tag(cfg, tag):
    """The cost_centers entry (dict: tag/code/name) matching this 'B<n>' tag
    string, or None if this property's config doesn't have one (or no config
    is loaded)."""
    if not cfg:
        return None
    for cc in cfg.get('cost_centers', []):
        if str(cc.get('tag')) == str(tag):
            return cc
    return None


def cost_center_for_name(cfg, name):
    """The cost_centers entry (dict: tag/code/name) whose display 'name'
    matches, e.g. resolving the free-text Building name box ('20 Riverside')
    back to its Kardin cost-center code ('west20') for a cross-bucket check
    that needs to filter another bucket's rows to just this building."""
    if not cfg or not name:
        return None
    for cc in cfg.get('cost_centers', []):
        if cc.get('name') == name:
            return cc
    return None


def committed_file_path(slug, filename):
    """Absolute path to a committed file under data/{slug}/, or None if it
    isn't there. Lets a property's assumption workbooks (GRP Budget
    Assumptions, Leasing Assumptions, 5-Year CapEx Plan, Interest
    Calculations, prior-year Expense Detail) be committed once per budget
    cycle - see config.yaml's optional `assumption_files` map - instead of
    re-uploaded by hand every session. filename is read from that map, e.g.
    cfg.get('assumption_files', {}).get('grp_budget_assumptions')."""
    if not slug or not filename:
        return None
    path = os.path.join(DATA_DIR, slug, filename)
    return path if os.path.isfile(path) else None


def generate_config_yaml(property_name, cost_centers):
    """cost_centers: [{'tag': str-or-None, 'code': str, 'name': str}, ...] -
    typically built from kardin_parser.parse_cost_center_roster (Kardin's own
    'Selected Cost Centers' report) plus user-assigned B<n> tags. Produces
    config.yaml text in the same style as the hand-written examples under
    data/ - ready to save as data/{slug}/config.yaml."""
    lines = [f"property_name: '{property_name}'", '', 'cost_centers:']
    for cc in cost_centers:
        tag = cc.get('tag')
        tag_str = f"'{tag}'" if tag else 'null'
        lines.append(f"  - tag: {tag_str}")
        lines.append(f"    code: '{cc['code']}'")
        lines.append(f"    name: '{cc['name']}'")
        lines.append('')
    return '\n'.join(lines).rstrip() + '\n'
