import io
import copy
import posixpath
import zipfile
from xml.etree import ElementTree as ET

import pytest
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

from app import validate_pptx
from merge import MergeError, merge_pptx


def rich_deck(text, color, wide=True):
    p = Presentation()
    p.slide_width, p.slide_height = (Inches(13.333), Inches(7.5)) if wide else (Inches(10), Inches(7.5))
    s = p.slides.add_slide(p.slide_layouts[5 if color == 'red' else 1])
    s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)).text = text
    image = io.BytesIO()
    Image.new('RGB', (20, 20), color).save(image, format='PNG')
    image.seek(0)
    s.shapes.add_picture(image, Inches(1), Inches(2), width=Inches(1))
    chart = CategoryChartData(); chart.categories = ['一', '二']; chart.add_series(text, [1, 3])
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(3), Inches(2), Inches(5), Inches(3), chart)
    s.notes_slide.notes_text_frame.text = text + ' 备注'
    b = io.BytesIO(); p.save(b)
    raw = b.getvalue()
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        entries = {n: z.read(n) for n in z.namelist()}
    theme = next(n for n in entries if n.startswith('ppt/theme/theme') and n.endswith('.xml'))
    entries[theme] = entries[theme].replace(b'Office Theme', f'{text} Theme'.encode())
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in entries.items(): z.writestr(n, data)
    return out.getvalue()


def check_relationships(raw):
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        names = set(z.namelist())
        for rels_name in [n for n in names if n.endswith('.rels')]:
            if rels_name == '_rels/.rels':
                source = ''
            else:
                directory, filename = posixpath.split(rels_name)
                source = posixpath.join(posixpath.dirname(directory), filename[:-5])
            for rel in ET.fromstring(z.read(rels_name)):
                if rel.attrib.get('TargetMode') != 'External':
                    target = posixpath.normpath(posixpath.join(posixpath.dirname(source), rel.attrib['Target']))
                    assert target in names, (rels_name, target)


def repack(raw, changes):
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        entries = {n: z.read(n) for n in z.namelist()}
    changes(entries)
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)
    return out.getvalue()


def test_real_merge_preserves_slides_media_charts_notes_and_themes():
    a, b = rich_deck('甲', 'red'), rich_deck('乙', 'blue')
    result = merge_pptx([a, b])
    assert validate_pptx(result) == 2
    check_relationships(result)
    p = Presentation(io.BytesIO(result))
    assert [[shape.text for shape in s.shapes if shape.has_text_frame and shape.text] for s in p.slides] == [['甲'], ['乙']]
    assert p.slides[0].slide_layout.name != p.slides[1].slide_layout.name
    for slide, label in zip(p.slides, ['甲', '乙']):
        assert any(shape.shape_type == 13 for shape in slide.shapes)
        assert any(shape.has_chart for shape in slide.shapes)
        assert label + ' 备注' in slide.notes_slide.notes_text_frame.text
    with zipfile.ZipFile(io.BytesIO(result)) as z:
        themes = [z.read(n) for n in z.namelist() if n.startswith('ppt/theme/') and n.endswith('.xml')]
        assert any(b'\xe7\x94\xb2 Theme' in t for t in themes)
        assert any(b'\xe4\xb9\x99 Theme' in t for t in themes)


def test_mixed_size_fails_without_partial_output():
    with pytest.raises(MergeError, match='页面比例'):
        merge_pptx([rich_deck('甲', 'red'), rich_deck('乙', 'blue', wide=False)])


def test_imported_ids_and_notes_master_cardinality():
    result = merge_pptx([rich_deck('甲', 'red'), rich_deck('乙', 'red')])
    with zipfile.ZipFile(io.BytesIO(result)) as z:
        root = ET.fromstring(z.read('ppt/presentation.xml'))
        ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
        masters = [int(x.get('id')) for x in root.findall('p:sldMasterIdLst/p:sldMasterId', ns)]
        slides = [int(x.get('id')) for x in root.findall('p:sldIdLst/p:sldId', ns)]
        assert len(masters) == len(set(masters)) and all(x >= 2147483648 for x in masters)
        assert len(slides) == len(set(slides)) and all(256 <= x < 2147483648 for x in slides)
        assert len(root.findall('p:notesMasterIdLst/p:notesMasterId', ns)) <= 1
        assert len([n for n in z.namelist() if n.startswith('ppt/notesMasters/notesMaster') and n.endswith('.xml')]) == 1
        types = ET.fromstring(z.read('[Content_Types].xml'))
        parts = [x.get('PartName') for x in types if x.tag.endswith('Override')]
        assert len(parts) == len(set(parts))


def test_presentation_namespace_and_order_survive_merge():
    first = rich_deck('甲', 'red')
    with zipfile.ZipFile(io.BytesIO(first)) as z:
        entries = {n: z.read(n) for n in z.namelist()}
    name = 'ppt/presentation.xml'
    entries[name] = entries[name].replace(b'<p:presentation ', b'<p:presentation xmlns:foo="urn:test:foo" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="foo" ', 1)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in entries.items(): z.writestr(n, data)
    result = merge_pptx([buf.getvalue(), rich_deck('乙', 'red')])
    with zipfile.ZipFile(io.BytesIO(result)) as z:
        xml = z.read(name)
        assert b'xmlns:foo="urn:test:foo"' in xml
        assert b'mc:Ignorable="foo"' in xml
        root = ET.fromstring(xml)
        tags = [x.tag.rsplit('}', 1)[-1] for x in root]
        assert tags.index('sldMasterIdLst') < tags.index('sldIdLst') < tags.index('sldSz')


def test_absolute_package_target_and_unreferenced_part():
    source = rich_deck('乙', 'red')
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        entries = {n: z.read(n) for n in z.namelist()}
    rel_name = 'ppt/slides/_rels/slide1.xml.rels'
    entries[rel_name] = entries[rel_name].replace(b'Target="../media/', b'Target="/ppt/media/')
    entries['ppt/unused.xml'] = b'<unused/>'
    entries['[Content_Types].xml'] = entries['[Content_Types].xml'].replace(
        b'</Types>', b'<Override PartName="/ppt/unused.xml" ContentType="application/xml"/></Types>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in entries.items(): z.writestr(n, data)
    result = merge_pptx([rich_deck('甲', 'red'), buf.getvalue()])
    with zipfile.ZipFile(io.BytesIO(result)) as z:
        assert 'ppt/joined1_unused.xml' not in z.namelist()
        assert b'/ppt/joined1_unused.xml' not in z.read('[Content_Types].xml')
    check_relationships(result)


def test_conflicting_notes_master_is_rejected():
    source = rich_deck('乙', 'red')
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        entries = {n: z.read(n) for n in z.namelist()}
    name = 'ppt/notesMasters/notesMaster1.xml'
    entries[name] = entries[name].replace(b'<p:notesMaster ', b'<p:notesMaster showMasterSp="0" ', 1)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in entries.items(): z.writestr(n, data)
    with pytest.raises(MergeError, match='统一备注母版'):
        merge_pptx([rich_deck('甲', 'red'), buf.getvalue()])


def test_opt_in_normalizes_conflicting_notes_masters_and_keeps_note_parts():
    first = rich_deck('甲', 'red')
    second = rich_deck('乙', 'red')
    def change(entries):
        name = 'ppt/notesMasters/notesMaster1.xml'
        entries[name] = entries[name].replace(b'<p:notesMaster ', b'<p:notesMaster showMasterSp="0" ', 1)
    second = repack(second, change)
    result = merge_pptx([first, second], normalize_notes_master=True)
    assert validate_pptx(result) == 2
    check_relationships(result)
    presentation = Presentation(io.BytesIO(result))
    assert len(presentation.slides) == 2
    assert [label in slide.notes_slide.notes_text_frame.text for label, slide in zip(['甲', '乙'], presentation.slides)] == [True, True]
    with zipfile.ZipFile(io.BytesIO(result)) as z:
        root = ET.fromstring(z.read('ppt/presentation.xml'))
        assert len([x for x in root if x.tag.endswith('notesMasterIdLst')]) == 1
        assert len([n for n in z.namelist() if n.startswith('ppt/notesMasters/') and n.endswith('.xml')]) == 1
        note_rels = [n for n in z.namelist() if n.startswith('ppt/notesSlides/_rels/')]
        assert len(note_rels) == 2
        master_targets = []
        for name in note_rels:
            rels = ET.fromstring(z.read(name))
            master_targets.extend(r.get('Target') for r in rels if r.get('Type', '').endswith('/notesMaster'))
        assert len(set(master_targets)) == 1


def test_opt_in_keeps_images_attached_to_both_notes_slides():
    decks = []
    for label, color in [('甲', 'red'), ('乙', 'blue')]:
        p = Presentation()
        slide = p.slides.add_slide(p.slide_layouts[5])
        slide.notes_slide.notes_text_frame.text = label + ' 备注'
        image = io.BytesIO()
        Image.new('RGB', (9, 9), color).save(image, format='PNG')
        slide.shapes.add_picture(io.BytesIO(image.getvalue()), Inches(1), Inches(1), width=Inches(1))
        out = io.BytesIO(); p.save(out)
        def add_note_image(entries):
            slide_xml = ET.fromstring(entries['ppt/slides/slide1.xml'])
            pic = next(x for x in slide_xml.iter() if x.tag.endswith('}pic'))
            embed = next(x for x in pic.iter() if x.tag.endswith('}blip'))
            embed.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', 'rId99')
            note_xml = ET.fromstring(entries['ppt/notesSlides/notesSlide1.xml'])
            tree = next(x for x in note_xml.iter() if x.tag.endswith('}spTree'))
            tree.append(copy.deepcopy(pic))
            entries['ppt/notesSlides/notesSlide1.xml'] = ET.tostring(note_xml)
            slide_rels = ET.fromstring(entries['ppt/slides/_rels/slide1.xml.rels'])
            image_rel = next(x for x in slide_rels if x.get('Type', '').endswith('/image'))
            note_rels = ET.fromstring(entries['ppt/notesSlides/_rels/notesSlide1.xml.rels'])
            note_image_rel = copy.deepcopy(image_rel)
            note_image_rel.set('Id', 'rId99')
            note_rels.append(note_image_rel)
            entries['ppt/notesSlides/_rels/notesSlide1.xml.rels'] = ET.tostring(note_rels)
        decks.append(repack(out.getvalue(), add_note_image))
    decks[1] = repack(decks[1], lambda e: e.__setitem__('ppt/notesMasters/notesMaster1.xml', e['ppt/notesMasters/notesMaster1.xml'].replace(b'<p:notesMaster ', b'<p:notesMaster showMasterSp="0" ', 1)))
    result = merge_pptx(decks, normalize_notes_master=True)
    check_relationships(result)
    p = Presentation(io.BytesIO(result))
    for slide, label in zip(p.slides, ['甲', '乙']):
        assert label in slide.notes_slide.notes_text_frame.text
        assert any(shape.shape_type == 13 for shape in slide.notes_slide.shapes)


def test_first_package_without_notes_master_imports_second_and_reuses_it():
    p = Presentation()
    p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
    p.slides.add_slide(p.slide_layouts[5])
    out = io.BytesIO(); p.save(out)
    first = out.getvalue()
    second = rich_deck('甲', 'red')
    third = rich_deck('乙', 'red')
    result = merge_pptx([first, second, third])
    check_relationships(result)
    p = Presentation(io.BytesIO(result))
    assert len(p.slides) == 3
    assert '甲 备注' in p.slides[1].notes_slide.notes_text_frame.text
    assert '乙 备注' in p.slides[2].notes_slide.notes_text_frame.text


def test_first_available_notes_master_is_used_when_first_deck_has_none():
    p = Presentation()
    p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
    p.slides.add_slide(p.slide_layouts[5])
    out = io.BytesIO(); p.save(out)
    second = rich_deck('甲', 'red')
    third = repack(rich_deck('乙', 'red'), lambda e: e.__setitem__(
        'ppt/notesMasters/notesMaster1.xml',
        e['ppt/notesMasters/notesMaster1.xml'].replace(b'<p:notesMaster ', b'<p:notesMaster showMasterSp="0" ', 1)))
    result = merge_pptx([out.getvalue(), second, third], normalize_notes_master=True)
    check_relationships(result)
    merged = Presentation(io.BytesIO(result))
    assert len(merged.slides) == 3
    assert '甲 备注' in merged.slides[1].notes_slide.notes_text_frame.text
    assert '乙 备注' in merged.slides[2].notes_slide.notes_text_frame.text
    with zipfile.ZipFile(io.BytesIO(result)) as z:
        assert len([n for n in z.namelist() if n.startswith('ppt/notesMasters/') and n.endswith('.xml')]) == 1


def test_conflicting_handout_master_stays_strict_with_notes_opt_in():
    def with_handout(raw, marker):
        def change(entries):
            part = 'ppt/handoutMasters/handoutMaster1.xml'
            entries[part] = f'<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" marker="{marker}"/>'.encode()
            rels_name = 'ppt/_rels/presentation.xml.rels'
            rels = ET.fromstring(entries[rels_name])
            ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
            ET.SubElement(rels, f'{{{ns}}}Relationship', Id='rId99', Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster', Target='handoutMasters/handoutMaster1.xml')
            entries[rels_name] = ET.tostring(rels)
            pres = ET.fromstring(entries['ppt/presentation.xml'])
            item = ET.Element('{http://schemas.openxmlformats.org/presentationml/2006/main}handoutMasterIdLst')
            ET.SubElement(item, '{http://schemas.openxmlformats.org/presentationml/2006/main}handoutMasterId').set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rId99')
            pres.insert(1, item)
            entries['ppt/presentation.xml'] = ET.tostring(pres)
            types = ET.fromstring(entries['[Content_Types].xml'])
            ET.SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Override', PartName='/' + part, ContentType='application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml')
            entries['[Content_Types].xml'] = ET.tostring(types)
        return repack(raw, change)
    with pytest.raises(MergeError, match='handoutMasterIdLst.*不兼容'):
        merge_pptx([with_handout(rich_deck('甲', 'red'), 'a'), with_handout(rich_deck('乙', 'red'), 'b')], normalize_notes_master=True)
