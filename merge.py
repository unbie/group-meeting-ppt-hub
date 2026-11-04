"""Merge PPTX packages by importing reachable OOXML part graphs."""
from __future__ import annotations

import copy
import io
import posixpath
import re
import zipfile
from lxml import etree as ET

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
Q = lambda ns, name: f'{{{ns}}}{name}'
LISTS = [('sldMasterIdLst', 'sldMasterId'), ('notesMasterIdLst', 'notesMasterId'), ('handoutMasterIdLst', 'handoutMasterId'), ('sldIdLst', 'sldId')]


class MergeError(ValueError):
    pass


def _xml(data):
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise MergeError('不支持包含 DTD 的 Office 文件')
    try:
        return ET.fromstring(data, parser=ET.XMLParser(resolve_entities=False, no_network=True))
    except ET.XMLSyntaxError as exc:
        raise MergeError('Office XML 无效') from exc


def _bytes(root):
    return ET.tostring(root, encoding='UTF-8', xml_declaration=True)


def _rels_path(part):
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')


def _target(source, target):
    if target.startswith('/'):
        result = posixpath.normpath(target.lstrip('/'))
    else:
        result = posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
    if result == '..' or result.startswith('../'):
        raise MergeError('关系目标超出 PPTX 包')
    return result


def _relations(pkg, part):
    name = _rels_path(part)
    return _xml(pkg[name]) if name in pkg else ET.Element(Q(PKG, 'Relationships'))


def _insert_list(pres, name):
    tag = Q(P, name)
    found = pres.find(tag)
    if found is not None:
        return found
    order = [n for n, _ in LISTS]
    new = ET.Element(tag)
    rank = order.index(name)
    for i, child in enumerate(pres):
        if child.tag in [Q(P, n) for n in order[rank + 1:]] or child.tag == Q(P, 'sldSz'):
            pres.insert(i, new)
            return new
    raise MergeError('演示文稿缺少可插入 ID 列表的位置')


def _same_graph(left_pkg, left_part, right_pkg, right_part, seen=None):
    """Compare single-instance master graphs, including their themes and media."""
    seen = seen or set()
    pair = (left_part, right_part)
    if pair in seen:
        return True
    seen.add(pair)
    if left_pkg[left_part] != right_pkg[right_part]:
        return False
    left_rels = _relations(left_pkg, left_part)
    right_rels = _relations(right_pkg, right_part)
    if len(left_rels) != len(right_rels):
        return False
    for a, b in zip(left_rels, right_rels):
        if any(a.get(k) != b.get(k) for k in ('Id', 'Type', 'TargetMode')):
            return False
        if a.get('TargetMode') == 'External':
            if a.get('Target') != b.get('Target'):
                return False
        elif not _same_graph(left_pkg, _target(left_part, a.get('Target')), right_pkg, _target(right_part, b.get('Target')), seen):
            return False
    return True


def merge_pptx(files: list[bytes], normalize_notes_master: bool = False) -> bytes:
    if not files:
        raise MergeError('该日期没有可合并的演示文稿')
    packages = []
    for raw in files:
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as z:
                names = z.namelist()
                if len(names) != len(set(names)):
                    raise MergeError('PPTX 包含重复部件')
                packages.append({n: z.read(n) for n in names})
        except (zipfile.BadZipFile, KeyError) as exc:
            raise MergeError('文件不是完整的 PPTX') from exc
    out = dict(packages[0])
    pres = _xml(out['ppt/presentation.xml'])
    rels = _xml(out['ppt/_rels/presentation.xml.rels'])
    ct = _xml(out['[Content_Types].xml'])
    size = pres.find(Q(P, 'sldSz'))
    if size is None:
        raise MergeError('演示文稿缺少页面尺寸')
    dimensions = (size.get('cx'), size.get('cy'))
    slide_list = _insert_list(pres, 'sldIdLst')
    slide_id = max([int(x.get('id')) for x in slide_list] + [255])
    master_ids = [int(x.get('id')) for x in pres.findall(f'{Q(P,"sldMasterIdLst")}/{Q(P,"sldMasterId")}')]
    for name in out:
        if re.fullmatch(r'ppt/slideMasters/[^/]+\.xml', name):
            master = _xml(out[name])
            master_ids += [int(x.get('id')) for x in master.findall(f'{Q(P,"sldLayoutIdLst")}/{Q(P,"sldLayoutId")}')]
    master_id = max(master_ids + [2147483647])
    rel_ids = {x.get('Id') for x in rels}
    next_rel = max([int(m.group(1)) for rid in rel_ids if (m := re.fullmatch(r'rId(\d+)', rid or ''))] + [0])
    defaults = {x.get('Extension').lower(): x.get('ContentType') for x in ct if x.tag == Q(CT, 'Default')}
    overrides = {x.get('PartName') for x in ct if x.tag == Q(CT, 'Override')}
    base_rels = {x.get('Id'): x for x in rels}
    base_singles = {}
    base_single_orig = {}
    singleton_types = {'notesMasterIdLst': 'notesMaster', 'handoutMasterIdLst': 'handoutMaster'}
    for list_name, item_name in LISTS[1:3]:
        item = pres.find(f'{Q(P,list_name)}/{Q(P,item_name)}')
        if item is not None:
            base_singles[list_name] = _target('ppt/presentation.xml', base_rels[item.get(Q(R, 'id'))].get('Target'))
        else:
            matches = [x for x in rels if x.get('Type', '').rsplit('/', 1)[-1] == singleton_types[list_name]]
            if len(matches) > 1:
                raise MergeError(f'{list_name} 包含多个母版关系')
            if matches:
                base_singles[list_name] = _target('ppt/presentation.xml', matches[0].get('Target'))
        if list_name in base_singles:
            base_single_orig[list_name] = (packages[0], base_singles[list_name])
    if normalize_notes_master and 'notesMasterIdLst' in base_singles and pres.find(Q(P, 'notesMasterIdLst')) is None:
        note_rel = next(x for x in rels if x.get('Type', '').endswith('/notesMaster'))
        ET.SubElement(_insert_list(pres, 'notesMasterIdLst'), Q(P, 'notesMasterId')).set(Q(R, 'id'), note_rel.get('Id'))

    for index, pkg in enumerate(packages[1:], 1):
        source_pres = _xml(pkg['ppt/presentation.xml'])
        source_size = source_pres.find(Q(P, 'sldSz'))
        if source_size is None or (source_size.get('cx'), source_size.get('cy')) != dimensions:
            raise MergeError('页面比例或尺寸不同；请先统一页面尺寸再合并')
        source_ct = _xml(pkg['[Content_Types].xml'])
        source_defaults = {x.get('Extension').lower(): x.get('ContentType') for x in source_ct if x.tag == Q(CT, 'Default')}
        source_overrides = {x.get('PartName').lstrip('/'): x.get('ContentType') for x in source_ct if x.tag == Q(CT, 'Override')}
        source_rels = {x.get('Id'): x for x in _xml(pkg['ppt/_rels/presentation.xml.rels'])}
        mapping = {'ppt/presentation.xml': 'ppt/presentation.xml'}
        single_replacements = {}
        for list_name, item_name in LISTS[1:3]:
            item = source_pres.find(f'{Q(P,list_name)}/{Q(P,item_name)}')
            matches = [x for x in source_rels.values() if x.get('Type', '').rsplit('/', 1)[-1] == singleton_types[list_name]]
            if len(matches) > 1:
                raise MergeError(f'{list_name} 包含多个母版关系')
            if not matches and item is None:
                continue
            rel = source_rels[item.get(Q(R, 'id'))] if item is not None else matches[0]
            old_part = _target('ppt/presentation.xml', rel.get('Target'))
            if list_name not in base_singles:
                # The source singleton will be imported through its list entry below.
                continue
            base_part = base_singles[list_name]
            original_pkg, original_part = base_single_orig[list_name]
            if not _same_graph(original_pkg, original_part, pkg, old_part):
                if list_name != 'notesMasterIdLst' or not normalize_notes_master:
                    if list_name == 'notesMasterIdLst':
                        raise MergeError('备注母版不兼容；如可接受备注排版变化，请勾选“统一备注母版（保留备注内容，备注排版可能变化）”')
                    raise MergeError(f'{list_name} 与首份演示文稿不兼容，无法无损合并')
            mapping[old_part] = base_part
            single_replacements[list_name] = old_part

        def import_part(name):
            nonlocal master_id
            if name in mapping:
                return mapping[name]
            if name not in pkg:
                raise MergeError(f'源文件关系缺失: {name}')
            directory, base = posixpath.split(name)
            new_name = posixpath.join(directory, f'joined{index}_{base}')
            if new_name in out:
                raise MergeError('合并部件名称冲突')
            mapping[name] = new_name
            data = pkg[name]
            if name.startswith('ppt/slideMasters/') and name.endswith('.xml'):
                root = _xml(data)
                for entry in root.findall(f'{Q(P,"sldLayoutIdLst")}/{Q(P,"sldLayoutId")}'):
                    master_id += 1
                    entry.set('id', str(master_id))
                data = _bytes(root)
            out[new_name] = data
            for ext, content_type in source_defaults.items():
                if ext not in defaults:
                    el = ET.Element(Q(CT, 'Default'), Extension=ext, ContentType=content_type)
                    ct.append(el)
                    defaults[ext] = content_type
                elif defaults[ext] != content_type:
                    raise MergeError(f'扩展名 {ext} 的内容类型冲突')
            if name in source_overrides:
                part_name = '/' + new_name
                if part_name in overrides:
                    raise MergeError('重复内容类型覆盖项')
                ct.append(ET.Element(Q(CT, 'Override'), PartName=part_name, ContentType=source_overrides[name]))
                overrides.add(part_name)
            old_rels_name = _rels_path(name)
            if old_rels_name in pkg:
                root = _xml(pkg[old_rels_name])
                for rel in root:
                    if rel.get('TargetMode') == 'External':
                        continue
                    old_target = _target(name, rel.get('Target'))
                    replacement = import_part(old_target)
                    rel.set('Target', posixpath.relpath(replacement, posixpath.dirname(new_name) or '.'))
                out[_rels_path(new_name)] = _bytes(root)
            return new_name

        for list_name, item_name in LISTS:
            source_list = source_pres.find(Q(P, list_name))
            if source_list is None and list_name in singleton_types and list_name not in single_replacements:
                matching = [x for x in source_rels.values() if x.get('Type', '').rsplit('/', 1)[-1] == singleton_types[list_name]]
                if matching:
                    source_list = ET.Element(Q(P, list_name))
                    ET.SubElement(source_list, Q(P, item_name)).set(Q(R, 'id'), matching[0].get('Id'))
            if source_list is None:
                continue
            if list_name in single_replacements:
                continue
            items = source_list.findall(Q(P, item_name))
            if list_name in ('notesMasterIdLst', 'handoutMasterIdLst') and (len(items) > 1 or list_name in base_singles and items):
                raise MergeError(f'{list_name} 数量超过规范限制')
            if not items:
                continue
            dest_list = _insert_list(pres, list_name)
            for item in items:
                old_rel = source_rels.get(item.get(Q(R, 'id')))
                if old_rel is None or old_rel.get('TargetMode') == 'External':
                    raise MergeError('演示文稿 ID 关系缺失或无效')
                old_part = _target('ppt/presentation.xml', old_rel.get('Target'))
                imported = import_part(old_part)
                next_rel += 1
                while f'rId{next_rel}' in rel_ids:
                    next_rel += 1
                new_rel = copy.deepcopy(old_rel)
                new_rel.set('Id', f'rId{next_rel}')
                new_rel.set('Target', posixpath.relpath(imported, 'ppt'))
                rels.append(new_rel)
                rel_ids.add(new_rel.get('Id'))
                new_item = copy.deepcopy(item)
                new_item.set(Q(R, 'id'), new_rel.get('Id'))
                if list_name == 'sldIdLst':
                    slide_id += 1
                    if slide_id >= 2147483648:
                        raise MergeError('幻灯片 ID 已用尽')
                    new_item.set('id', str(slide_id))
                elif list_name == 'sldMasterIdLst':
                    master_id += 1
                    if master_id > 4294967295:
                        raise MergeError('母版 ID 已用尽')
                    new_item.set('id', str(master_id))
                dest_list.append(new_item)
                if list_name in ('notesMasterIdLst', 'handoutMasterIdLst'):
                    base_singles[list_name] = imported
                    base_single_orig[list_name] = (pkg, old_part)

    out['ppt/presentation.xml'] = _bytes(pres)
    out['ppt/_rels/presentation.xml.rels'] = _bytes(rels)
    out['[Content_Types].xml'] = _bytes(ct)
    data = io.BytesIO()
    with zipfile.ZipFile(data, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, content in out.items():
            z.writestr(name, content)
    return data.getvalue()
