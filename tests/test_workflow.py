"""Synthetic technical fixtures only: no research facts or real user approvals."""
import copy
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
import project
import office


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(prefix='research-wh-technical-test-')
        self.work=Path(self.temp.name)
        self.state={'mode':'new','skill':{**project.read(ROOT/'VERSION.json'),'sha256':project.package_hash()},'approvals':{}}
        self.write('source.txt','Synthetic technical input, not a paper.')
        self.json('intake.json',{'duration_minutes':5,'background':'technical fixture',
            'documents':[{'id':'p','path':'source.txt','sha256':project.sha(self.work/'source.txt'),
                          'chapters':[{'id':'c','title':'synthetic section'}]}]})
        self.json('figures.json',{'selected_assets':[]})
        self.write('content.md','Synthetic content selection.')
        self.write('overview.md','Synthetic overview.')
        self.write('chapter.md','Synthetic chapter.')
        self.pages=[{'id':f's{i}','title':f'Technical title {i}','header':'1. Technical fixture',
                    'on_screen':[f'Technical body {i}'],'assets':[]} for i in range(1,4)]
        self.json('outline.json',{'pages':self.pages})
        self.json('selection.json',{'pages':['s1','s2','s3']})
        self.json('style.json',project.read(ROOT/'assets/style.json'))
        self.plan={'outline_file':'outline.json','style_file':'style.json','assets':[],
                   'permissions':[],'pages':[{'id':p['id'],'title':p['title'],'header':p['header'],
                    'notes':'Synthetic detailed notes; no real research statements.',
                    'objects':[{'id':'body','type':'text','text':p['on_screen'][0],
                                'box':[24,120,900,300],'bold':False}]} for p in self.pages]}
        self.json('plan.json',self.plan)
        for scope,file in [('intake','intake.json'),('figures','figures.json'),('content','content.md'),
            ('reading-overview:p','overview.md'),('reading-chapter:p:c','chapter.md'),
            ('outline','outline.json'),('prototype-selection','selection.json')]:
            self.approve(scope,file)

    def tearDown(self):
        self.temp.cleanup()

    def write(self,name,text):
        (self.work/name).write_text(text,encoding='utf-8')

    def json(self,name,obj):
        project.save(self.work/name,obj)

    def approve(self,scope,name):
        p=self.work/name
        self.state['approvals'][scope]={'path':str(p),'sha256':project.sha(p),
          'quote':'SYNTHETIC TEST ONLY - not a user approval'}

    def check(self,stage='prototype'):
        return project.check_plan(self.state,self.work,self.work/'plan.json',stage)

    def test_prototype_valid_records(self):
        self.assertEqual(self.check(),['s1','s2','s3'])

    def test_full_blocked_before_prototype_confirmation(self):
        with self.assertRaisesRegex(ValueError,'prototypes'):
            self.check('full')

    def test_source_change_invalidates_reading(self):
        self.write('source.txt','changed source')
        with self.assertRaisesRegex(ValueError,'原文'):
            self.check()

    def test_chapter_confirmation_required(self):
        del self.state['approvals']['reading-chapter:p:c']
        with self.assertRaisesRegex(ValueError,'reading-chapter'):
            self.check()

    def test_outline_change_invalidates_confirmation(self):
        self.json('outline.json',{'pages':self.pages+[],'changed':True})
        with self.assertRaisesRegex(ValueError,'失效'):
            self.check()

    def test_layout_only_change_does_not_invalidate_content(self):
        self.plan['pages'][0]['objects'][0]['box']=[40,150,800,230]
        self.json('plan.json',self.plan)
        self.assertEqual(len(self.check()),3)

    def test_unapproved_text_is_blocked(self):
        self.plan['pages'][0]['objects'][0]['text']='Unapproved scientific claim'
        self.json('plan.json',self.plan)
        with self.assertRaisesRegex(ValueError,'未经大纲确认'):
            self.check()

    def test_all_pages_cannot_be_smuggled_as_prototypes(self):
        self.plan['pages'].append({**self.plan['pages'][0],'id':'extra'})
        self.json('plan.json',self.plan)
        with self.assertRaisesRegex(ValueError,'代表页'):
            self.check()

    def test_changed_snapshot_attachment_is_blocked(self):
        self.write('prototype.pptx','synthetic attachment')
        self.json('review.json',{'files':[{'path':'prototype.pptx','sha256':project.sha(self.work/'prototype.pptx')},
                                        {'path':'style.json','sha256':project.sha(self.work/'style.json')}]})
        self.approve('prototypes','review.json')
        self.assertEqual(len(self.check('full')),3)
        self.write('prototype.pptx','changed synthetic attachment')
        with self.assertRaisesRegex(ValueError,'附件'):
            self.check('full')

    def test_style_change_after_prototypes_blocked(self):
        self.json('review.json',{'files':[{'path':'style.json','sha256':project.sha(self.work/'style.json')}]})
        self.approve('prototypes','review.json')
        style=project.read(self.work/'style.json');style['body_font_pt']+=1
        self.json('style.json',style)
        with self.assertRaisesRegex(ValueError,'附件'):
            self.check('full')

    def test_changed_skill_not_silently_used(self):
        self.state['skill']['sha256']='not-current-version'
        with self.assertRaisesRegex(ValueError,'skill'):
            self.check()

    def test_unapproved_permission_blocked(self):
        self.write('crop.json','synthetic proposal')
        self.plan['permissions']=[{'scope':'transform:fig1','proposal':'crop.json'}]
        self.json('plan.json',self.plan)
        with self.assertRaisesRegex(ValueError,'transform:fig1'):
            self.check()

    def add_image(self,origin='source'):
        self.write('image.png','SYNTHETIC hash-only image fixture, not actual artwork')
        digest=project.sha(self.work/'image.png')
        self.json('figures.json',{'selected_assets':['f'], 'files':[{'path':'image.png','sha256':digest}]})
        self.approve('figures','figures.json')
        self.pages[0]['assets']=['f'];self.json('outline.json',{'pages':self.pages});self.approve('outline','outline.json')
        self.plan['assets']=[{'id':'f','path':'image.png','sha256':digest,'origin':origin}]
        self.plan['pages'][0]['objects'].append({'id':'image','type':'image','asset_id':'f','box':[600,180,200,100]})
        self.json('plan.json',self.plan)

    def test_bound_selected_original_allowed(self):
        self.add_image();self.assertEqual(len(self.check()),3)

    def test_changed_image_cannot_be_hidden_by_new_plan_hash(self):
        self.add_image();self.write('image.png','changed synthetic image')
        self.plan['assets'][0]['sha256']=project.sha(self.work/'image.png');self.json('plan.json',self.plan)
        with self.assertRaisesRegex(ValueError,'审核附件'):
            self.check()

    def test_external_image_requires_specific_permission(self):
        self.add_image('external')
        with self.assertRaisesRegex(ValueError,'add-image:f'):
            self.check()

    def test_implicit_cover_crop_is_blocked(self):
        self.add_image();self.plan['pages'][0]['objects'][-1]['fit']='cover';self.json('plan.json',self.plan)
        with self.assertRaisesRegex(ValueError,'裁剪'):
            self.check()

    def test_font_replacement_requires_confirmation(self):
        style=project.read(self.work/'style.json');style['font']='UnapprovedFont';self.json('style.json',style)
        with self.assertRaisesRegex(ValueError,'字体替换'):
            self.check()

    def test_archive_never_overwrites_prior_revision(self):
        self.json('bundle.json',{})
        (self.work/'deliver'/'v01').mkdir(parents=True)
        with self.assertRaisesRegex(ValueError,'拒绝覆盖'):
            project.archive(self.state,self.work,self.work/'bundle.json',self.work/'deliver','v01')

    def test_word_line_breaks_and_tabs_are_not_lost(self):
        doc=self.work/'breaks.docx'
        with zipfile.ZipFile(doc,'w') as z:
            z.writestr('word/document.xml','<w:document xmlns:w="'+office.NS['w']+'"><w:body><w:p><w:r><w:t>First</w:t><w:br/><w:t>Second</w:t><w:tab/><w:t>Third</w:t></w:r></w:p></w:body></w:document>')
        self.assertEqual(office.docx_paragraphs(doc),['First\nSecond\tThird'])

    def prepare_edit(self):
        self.state['mode']='edit'
        self.write('edit-scope.md','Synthetic edit scope')
        self.write('edit-reading.md','Synthetic edit reading')
        for scope in ['edit-scope','edit-reading']:
            self.approve(scope,scope+'.md')

    def test_edit_does_not_bypass_image_permission(self):
        self.prepare_edit();self.write('crop.json','Synthetic crop proposal')
        self.json('edit-plan.json',{'permissions':[{'scope':'transform:f','proposal':'crop.json'}]})
        self.approve('edit-plan','edit-plan.json')
        with self.assertRaisesRegex(ValueError,'transform:f'):
            project.check_plan(self.state,self.work,self.work/'unused.json','edit')

    def test_edit_source_snapshot_change_is_blocked(self):
        self.prepare_edit()
        self.json('edit-plan.json',{'files':[{'path':'source.txt','sha256':project.sha(self.work/'source.txt')}]})
        self.approve('edit-plan','edit-plan.json');self.write('source.txt','changed')
        with self.assertRaisesRegex(ValueError,'附件'):
            project.check_plan(self.state,self.work,self.work/'unused.json','edit')

    def test_init_rejects_task_inside_skill(self):
        r=subprocess.run([sys.executable,str(ROOT/'scripts/project.py'),'init','--workdir',str(ROOT/'work')],capture_output=True)
        self.assertEqual(r.returncode,2)
        self.assertFalse((ROOT/'work'/'project.json').exists())


if __name__=='__main__':
    unittest.main()
