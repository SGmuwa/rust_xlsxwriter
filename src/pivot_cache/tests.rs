// Pivot cache unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod pivot_cache_tests {

    use crate::pivot_cache::{PivotCache, PivotFieldValue};
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

    #[test]
    fn test_assemble_shared_items() {
        // The distinct values of the fields that a pivot table uses on an
        // axis. The last field isn't used by a pivot table and is left empty.
        let mut cache = PivotCache::new("Data", 0, 0, 4, 3);
        cache.field_names = ["Region", "Volume", "Month", "Note"]
            .iter()
            .map(ToString::to_string)
            .collect();

        cache.field_values = vec![
            vec![
                PivotFieldValue::String("East".to_string()),
                PivotFieldValue::String("West".to_string()),
            ],
            vec![
                PivotFieldValue::Number(1000.0),
                PivotFieldValue::Number(9000.0),
            ],
            vec![PivotFieldValue::Number(1.5), PivotFieldValue::Blank],
            vec![],
        ];

        cache.assemble_xml_file();

        let got = xmlwriter::cursor_to_str(&cache.writer);
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" saveData="0" refreshOnLoad="1" refreshedBy="rust_xlsxwriter" refreshedDate="0" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="4">
                    <cacheSource type="worksheet">
                        <worksheetSource ref="A1:D5" sheet="Data"/>
                    </cacheSource>
                    <cacheFields count="4">
                        <cacheField name="Region" numFmtId="0">
                            <sharedItems count="2"><s v="East"/><s v="West"/></sharedItems>
                        </cacheField>
                        <cacheField name="Volume" numFmtId="0">
                            <sharedItems count="2" containsNumber="1" containsString="0" containsSemiMixedTypes="0" minValue="1000" maxValue="9000"><n v="1000"/><n v="9000"/></sharedItems>
                        </cacheField>
                        <cacheField name="Month" numFmtId="0">
                            <sharedItems count="2" containsBlank="1" containsNumber="1" containsString="0" minValue="1.5" maxValue="1.5"><n v="1.5"/><m/></sharedItems>
                        </cacheField>
                        <cacheField name="Note" numFmtId="0"><sharedItems/></cacheField>
                    </cacheFields>
                </pivotCacheDefinition>
            "#,
        );

        assert_eq!(expected, got);
    }
}
