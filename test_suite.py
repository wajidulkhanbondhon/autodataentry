import os
import unittest
import app

class TestSmartAdmissionBot(unittest.TestCase):
    def test_clean_data_phone(self):
        val, mod, reason = app.AIGuardianEngine.clean_data('01711123456.0', 'phone')
        self.assertEqual(val, '01711123456')
        self.assertTrue(mod)

        # with +880
        val, mod, reason = app.AIGuardianEngine.clean_data('+8801711123456', 'phone')
        self.assertEqual(val, '01711123456')

        # missing leading 0 (10 digits starting with 1)
        val, mod, reason = app.AIGuardianEngine.clean_data('1711123456', 'phone')
        self.assertEqual(val, '01711123456')

    def test_optimize_steps_list(self):
        steps = [
            {'action': 'click', 'selector': '#input1'},
            {'action': 'type', 'selector': '#input1', 'value': 'test'},
            {'action': 'click', 'selector': '#submit'},
            {'action': 'click', 'selector': '#submit'}
        ]
        cleaned, removed = app.optimize_steps_list(steps)
        self.assertEqual(removed, 2)
        self.assertEqual(len(cleaned), 2)
        self.assertEqual(cleaned[0]['action'], 'type')
        self.assertEqual(cleaned[1]['action'], 'click')

    def test_preflight_audit(self):
        rows = [
            {'Phone': '01711123456', 'Reg': '101'},
            {'Phone': '01711123456', 'Reg': '101'}
        ]
        steps = [
            {'col_bind': 'Phone', 'action': 'type'},
            {'col_bind': 'Reg', 'action': 'type'}
        ]
        audit = app.AIGuardianEngine.preflight_audit(rows, steps)
        self.assertEqual(audit['status'], 'ISSUES_FOUND')
        self.assertGreaterEqual(audit['issues_count'], 1)

    def test_detect_headers_and_normalize_rows(self):
        # Headerless student data rows
        raw_rows = [
            (1602, 4956750, 'Shahrin', '8801739407522', '-', '-', 'RAJSHAHI GOVT GIRLS HIGH SCHOOL', 'Female', 'Islam'),
            (1603, 4956790, 'Hamima', '8801772937474', '-', '-', 'RAJSHAHI GOVT GIRLS HIGH SCHOOL', 'Female', 'Islam'),
        ]
        headers, rows, is_hl = app.detect_headers_and_normalize_rows(raw_rows)
        self.assertTrue(is_hl)
        self.assertIn('Roll_No', headers)
        self.assertIn('Phone', headers)
        self.assertIn('Institute', headers)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]['Roll_No'], '4956750')
        self.assertEqual(rows[1]['Roll_No'], '4956790')

    def test_extract_clean_col_semantic(self):
        row = {
            'Roll_No': '4956750',
            'Name': 'Fatema Akter',
            'Institute': 'RAJSHAHI GOVT GIRLS HIGH SCHOOL',
            'Phone': '01739407522',
            'Entry_Status': ''
        }
        # 'Reg Number' bound step matching 'Roll_No' via semantic group
        self.assertEqual(app.extract_clean_col(row, 'Reg Number'), '4956750')
        self.assertEqual(app.extract_clean_col(row, 'Roll'), '4956750')
        self.assertEqual(app.extract_clean_col(row, 'registration'), '4956750')
        # 'Institute' exact and semantic
        self.assertEqual(app.extract_clean_col(row, 'Institute'), 'RAJSHAHI GOVT GIRLS HIGH SCHOOL')
        self.assertEqual(app.extract_clean_col(row, 'School Name'), 'RAJSHAHI GOVT GIRLS HIGH SCHOOL')
        # Mobile / Phone
        self.assertEqual(app.extract_clean_col(row, 'Mobile'), '01739407522')
        self.assertEqual(app.extract_clean_col(row, 'Contact No'), '01739407522')
        # Name
        self.assertEqual(app.extract_clean_col(row, 'Student Name'), 'Fatema Akter')
        # Empty/non-existent
        self.assertEqual(app.extract_clean_col(row, 'NonExistentCol'), '')
        self.assertEqual(app.extract_clean_col({}, 'Roll_No'), '')

    def test_recorded_steps_compatibility_with_real_excel(self):
        excel_path = r'c:\Users\Googe\Desktop\Class-9 Math olympiad Data for Data Entry - Copy.xlsx'
        if os.path.exists(excel_path):
            headers, rows, is_hl = app.universal_read_file(excel_path)
            self.assertGreater(len(rows), 0)
            sample_row = rows[0]
            
            # Real recorded steps bindings
            steps = app.safe_json_load(app.STEPS_FILE, [])
            for step in steps:
                bind = step.get('col_bind', '').strip()
                if bind:
                    extracted = app.extract_clean_col(sample_row, bind)
                    # Verify crucial bindings like roll or institute extract correctly from real excel
                    if any(k in bind.lower() for k in ['roll', 'reg']):
                        self.assertTrue(len(extracted) >= 5, f"Roll extraction failed for binding: {bind}")
                    if 'institute' in bind.lower():
                        self.assertTrue(len(extracted) >= 3, f"Institute extraction failed for binding: {bind}")

    def test_secure_settings_save_and_load(self):
        test_path = 'test_tmp_settings.json'
        try:
            sample_settings = {
                'portal_username': 'test_user',
                'portal_password': 'mySuperSecretPassword!#@1',
                'speed_ms': 900
            }
            app.safe_json_save(test_path, sample_settings)
            
            # verify on-disk contents
            with open(test_path, 'r', encoding='utf-8') as f:
                content = f.read()
            self.assertNotIn('mySuperSecretPassword!#@1', content)
            self.assertIn('_enc_portal_password', content)

            # verify loading
            loaded = app.safe_json_load(test_path, {})
            self.assertEqual(loaded.get('portal_password'), 'mySuperSecretPassword!#@1')
            self.assertEqual(loaded.get('portal_username'), 'test_user')
        finally:
            if os.path.exists(test_path):
                os.remove(test_path)

if __name__ == '__main__':
    unittest.main()

