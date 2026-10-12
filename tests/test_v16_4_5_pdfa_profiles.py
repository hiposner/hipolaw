"""Regression: Word's PDF/A-3a XMP must not be rejected as 'not PDF/A-1'.

Synthetic PDF metadata is deliberately not treated as proof of compliance;
unit tests cover profile identification and safe validation plumbing only.
"""
import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('hipolaw_pdfa', ROOT/'generate_legal_tables_v16_4_8.py')
mod=importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


def xmp(part='3', level='A', *, attributes=False):
    ns='http://www.aiim.org/pdfa/ns/id/'
    if attributes:
        body=f'<rdf:Description rdf:about="" xmlns:pdfaid="{ns}" pdfaid:part="{part}" pdfaid:conformance="{level}" />'
    else:
        body=(f'<rdf:Description rdf:about="" xmlns:pdfaid="{ns}">'
              f'<pdfaid:part>{part}</pdfaid:part><pdfaid:conformance>{level}</pdfaid:conformance>'
              f'</rdf:Description>')
    return ('<x:xmpmeta xmlns:x="adobe:ns:meta/">'
            '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
            +body+'</rdf:RDF></x:xmpmeta>')


@unittest.skipIf(mod.pymupdf is None, 'PyMuPDF not installed')
class PdfaProfileCompatibility(unittest.TestCase):
    def _make_pdf(self, folder, *, metadata=None):
        pdf=Path(folder)/'generic_motion.pdf'
        doc=mod.pymupdf.open()
        page=doc.new_page()
        page.insert_text((60,70),'Generic legal document - introduction',fontsize=12)
        if metadata is not None:
            doc.set_xml_metadata(metadata)
        doc.save(pdf)
        doc.close()
        return pdf

    def test_supported_word_pdfa_profiles(self):
        for version,level in [('1','A'),('1','B'),('2','A'),('2','B'),('2','U'),
                              ('3','A'),('3','B'),('3','U')]:
            with self.subTest(profile=version+level):
                self.assertEqual(mod.recognized_pdfa_profile(xmp(version,level)),version+level.lower())

    def test_rdf_attributes_are_supported(self):
        self.assertEqual(mod.recognized_pdfa_profile(xmp('3','A',attributes=True)),'3a')

    def test_xmp_missing_or_conflicting_declaration_is_rejected(self):
        for metadata in ('', '<x:xmpmeta xmlns:x="adobe:ns:meta/"/>',
                         '<x:xmpmeta', xmp('4','A'), xmp('1','U'),
                         xmp('3','A').replace('</rdf:RDF>', '<rdf:Description xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"><pdfaid:part>2</pdfaid:part></rdf:Description></rdf:RDF>')):
            with self.subTest(xmp=metadata[:45]):
                with self.assertRaisesRegex(RuntimeError,'FILING BLOCKED'):
                    mod.recognized_pdfa_profile(metadata)

    def test_pdfa3a_is_accepted_for_audit_with_warning_without_verapdf(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf=self._make_pdf(tmp,metadata=xmp('3','A'))
            # Synthetic XMP only; audit must report incomplete ISO conformance evidence.
            with patch.object(mod.shutil,'which',return_value=None):
                with patch.dict(mod.os.environ,{'VERAPDF_CLI':''}):
                    results,warnings=mod.pdfa_safety_audit(pdf)
            self.assertIn('PDF/A declaration in XMP: PDF/A-3a',results)
            self.assertTrue(any('not independently verified' in w for w in warnings))
            self.assertTrue(any('NOT RUN' in x for x in results))

    def test_non_pdfa_is_still_filing_blocked(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf=self._make_pdf(tmp)
            with self.assertRaisesRegex(RuntimeError,'FILING BLOCKED: PDF/A identification metadata is missing'):
                mod.pdfa_safety_audit(pdf)

    def test_independent_verapdf_receives_detected_profile(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf=self._make_pdf(tmp,metadata=xmp('3','A'))
            success=type('Completed',(),{'returncode':0,'stdout':'<validationReport isCompliant="true"/>','stderr':''})()
            with patch.object(mod.subprocess,'run',return_value=success) as run:
                results,warnings=mod.pdfa_safety_audit(pdf,independent_validator='verapdf_mock')
            args=run.call_args.args[0]
            self.assertEqual(args[:3],['verapdf_mock','-f','3a'])
            self.assertIn('Independent veraPDF PDF/A-3a validation: PASS',results)
            self.assertFalse(any('not independently verified' in w for w in warnings))

    def test_verapdf_failure_still_blocks(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf=self._make_pdf(tmp,metadata=xmp('3','A'))
            failure=type('Completed',(),{'returncode':0,'stdout':'<validationReport isCompliant="false"/>','stderr':''})()
            with patch.object(mod.subprocess,'run',return_value=failure):
                with self.assertRaisesRegex(RuntimeError,'FILING BLOCKED: independent veraPDF PDF/A-3a'):
                    mod.pdfa_safety_audit(pdf,independent_validator='verapdf_mock')


if __name__=='__main__':
    unittest.main()
