// Pivot cache unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod pivot_cache_tests {

    use crate::pivot_cache::PivotCache;
    use crate::test_functions::xml_to_vec;
    use crate::xmlwriter;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble1() {
        let mut cache = PivotCache::new("Data", 0, 0, 50, 3);
        cache.field_names = ["Region", "Item", "Volume", "Month"]
            .iter()
            .map(ToString::to_string)
            .collect();

        cache.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&cache.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" saveData="0" refreshOnLoad="1" refreshedBy="rust_xlsxwriter" refreshedDate="0" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="50">
                    <cacheSource type="worksheet">
                        <worksheetSource ref="A1:D51" sheet="Data"/>
                    </cacheSource>
                    <cacheFields count="4">
                        <cacheField name="Region" numFmtId="0"><sharedItems/></cacheField>
                        <cacheField name="Item" numFmtId="0"><sharedItems/></cacheField>
                        <cacheField name="Volume" numFmtId="0"><sharedItems/></cacheField>
                        <cacheField name="Month" numFmtId="0"><sharedItems/></cacheField>
                    </cacheFields>
                </pivotCacheDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }
}
