"""Synthetic cache fixtures only; no real user approval or visual inspection."""
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
import render_cache as cache
import office
import project

P=office.NS['p']; R=office.NS['r']; REL=office.NS['rel']


class CacheTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(prefix='research-wh-cache-fixture-')
        self.w=Path(self.tmp.name); self.ppt=self.w/'source.pptx';self.c=self.w/'cache';self.n=0
        self.renderer={'application':'PowerPoint','version':'synthetic','width':1600,'height':900,'fonts':'synthetic'}
        self.parts={
          'ppt/presentation.xml':f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="1" r:id="s1"/><p:sldId id="2" r:id="s2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
          'ppt/_rels/presentation.xml.rels':self.rs([('s1','slide','slides/slide1.xml'),('s2','slide','slides/slide2.xml')]),
          'ppt/slides/slide1.xml':'<slide>fixture one</slide>',
          'ppt/slides/slide2.xml':'<slide>fixture two</slide>',
          'ppt/slides/_rels/slide1.xml.rels':self.rs([('n','notesSlide','../notesSlides/n.xml'),('i','image','../media/i.png'),('m','slideLayout','../slideLayouts/l.xml')]),
          'ppt/slides/_rels/slide2.xml.rels':self.rs([('m','slideLayout','../slideLayouts/l.xml')]),
          'ppt/slideLayouts/l.xml':'<layout>fixture</layout>',
          'ppt/slideLayouts/_rels/l.xml.rels':self.rs([('t','theme','../theme/t.xml')]),
          'ppt/theme/t.xml':'<theme>fixture theme</theme>',
          'ppt/notesSlides/n.xml':'<notes>synthetic notes</notes>',
          'ppt/media/i.png':b'\x89PNG\r\n\x1a\nSYNTHETIC-FIXTURE'}
        self.save()

    def rs(self,items):
        return '<Relationships xmlns="'+REL+'">'+''.join(f'<Relationship Id="{i}" Type="{R}/{k}" Target="{t}"/>' for i,k,t in items)+'</Relationships>'

    def save(self):
        with zipfile.ZipFile(self.ppt,'w') as z:
            for k,v in self.parts.items():z.writestr(k,v)

    def plan(self,renderer=None,pages=None):
        self.n+=1; out=self.w/f'out{self.n}'
        return cache.plan(self.ppt,out,self.c,renderer or self.renderer,pages),out

    def seed(self):
        p,out=self.plan()
        rows=[]
        for j in p['jobs']:
            f=out/j['file'];f.write_bytes(b'\x89PNG\r\n\x1a\nSYNTHETIC-PREVIEW'+str(j['page']).encode())
            rows.append({**j,'sha256':office.sha(f)})
        r={**p,'opened':True,'slides':rows}
        report=out/'render-report.json';cache.write(report,r);cache.complete(report)
        return report

    def tearDown(self):self.tmp.cleanup()

    def test_notes_only_change_reuses_all_slides(self):
        self.seed();self.parts['ppt/notesSlides/n.xml']='<notes>changed notes</notes>';self.save()
        p,_=self.plan();self.assertTrue(all(j['cache_hit'] for j in p['jobs']))

    def test_relationship_id_rewrite_keeps_same_visual_fingerprint(self):
        before=cache.fingerprints(self.ppt)
        self.parts['ppt/presentation.xml']=self.parts['ppt/presentation.xml'].replace('r:id="s1"','r:id="new1"')
        self.parts['ppt/_rels/presentation.xml.rels']=self.parts['ppt/_rels/presentation.xml.rels'].replace('Id="s1"','Id="new1"')
        self.parts['ppt/slides/_rels/slide1.xml.rels']=self.parts['ppt/slides/_rels/slide1.xml.rels'].replace('Id="m"','Id="newm"')
        self.save();self.assertEqual(before,cache.fingerprints(self.ppt))

    def test_nonvisual_creation_id_does_not_force_new_preview(self):
        self.parts['ppt/slides/slide1.xml']='<s xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"><a16:creationId id="one"/><text>same</text></s>'
        self.save();before=cache.fingerprints(self.ppt)
        self.parts['ppt/slides/slide1.xml']=self.parts['ppt/slides/slide1.xml'].replace('id="one"','id="two"')
        self.save();self.assertEqual(before,cache.fingerprints(self.ppt))

    def test_one_slide_change_invalidates_only_that_page(self):
        self.seed();self.parts['ppt/slides/slide1.xml']='<slide>changed</slide>';self.save()
        p,_=self.plan();self.assertEqual([j['cache_hit'] for j in p['jobs']],[False,True])

    def test_shared_theme_change_invalidates_both_pages(self):
        self.seed();self.parts['ppt/theme/t.xml']='<theme>changed</theme>';self.save()
        p,_=self.plan();self.assertFalse(any(j['cache_hit'] for j in p['jobs']))

    def test_original_image_change_invalidates_its_page(self):
        self.seed();self.parts['ppt/media/i.png']=b'changed image';self.save()
        p,_=self.plan();self.assertEqual([j['cache_hit'] for j in p['jobs']],[False,True])

    def test_renderer_and_font_changes_invalidate_cache(self):
        self.seed();p,_=self.plan({**self.renderer,'fonts':'different'})
        self.assertFalse(any(j['cache_hit'] for j in p['jobs']))

    def test_corrupt_cached_png_is_not_reused(self):
        report=self.seed();key=cache.read(report)['slides'][0]['cache_key'];(self.c/(key+'.png')).write_bytes(b'corrupt')
        p,_=self.plan();self.assertEqual([j['cache_hit'] for j in p['jobs']],[False,True])

    def test_partial_export_does_not_claim_complete_deck(self):
        p,_=self.plan(pages=[2]);self.assertFalse(p['complete_deck']);self.assertEqual([j['page'] for j in p['jobs']],[2])

    def test_cache_hit_does_not_invent_visual_review(self):
        self.seed();p,_=self.plan();self.assertFalse(any(j['visual_reviewed'] for j in p['jobs']))

    def test_existing_review_can_be_reused(self):
        r=self.seed();cache.review(r,[1],'SYNTHETIC TEST ONLY; not actual human inspection')
        p,_=self.plan();self.assertEqual([j['visual_review_reused'] for j in p['jobs']],[True,False])

    def test_powerpoint_only_bundle_can_archive_without_wps(self):
        # Use the shipped empty template and a corresponding real OOXML notes document.
        ppt=ROOT/'assets/reference-layouts.pptx'; slides=office.ppt_slides(ppt)
        doc=self.w/'notes.docx';office.notes_docx(ppt,doc)
        preview=self.w/'previews';preview.mkdir()
        rows=[]
        for i in range(1,len(slides)+1):
            f=preview/f'slide-{i:03}.png';f.write_bytes(b'SYNTHETIC PREVIEW')
            rows.append({'page':i,'file':f.name,'sha256':office.sha(f)})
        sh=office.sha(ppt)
        cache.write(preview/'render-report.json',{'source_sha256':sh,'application':'PowerPoint','opened':True,'slides':rows})
        thumb=self.w/'contact.png';thumb.write_bytes(b'SYNTHETIC CONTACT')
        cache.write(str(thumb)+'.source.json',{'source_sha256':sh,'contact_sheet_sha256':office.sha(thumb)})
        qa=self.w/'qa.json';cache.write(qa,{'source_sha256':sh,'visual_reviewed':True,'content_reviewed':True,'unresolved_errors':[],'note':'SYNTHETIC TEST ONLY'})
        manifest=self.w/'bundle.json';cache.write(manifest,{'pptx':str(ppt),'docx':str(doc),'preview_dir':str(preview),'contact_sheet':str(thumb),'qa':str(qa),'software_reports':{'PowerPoint':str(preview/'render-report.json')},'formal_pages':len(slides)})
        dest=project.archive({'skill':{'version':'synthetic'}},self.w,manifest,self.w/'deliver','v01')
        self.assertTrue((Path(dest)/'version-record.json').is_file())


if __name__=='__main__':unittest.main()
