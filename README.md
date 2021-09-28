# Excel to JSON converter
### Proof of concept, TDD using integration tests
#### Resources
1. Apache Poi library - "Don't reinvent the wheel"
2. Use an example - https://www.dev2qa.com/convert-excel-to-json-in-java-example/
3. Java 11 - What was on hand
4. NOTE: Above example used ancient, outdated JSON lib. Direct swap to org.json implementation 
5. NOTE: This was still done "TDD" - it's just an integration test.  I don't want to know how Poi's guts work!
6. Sample Excel file direct from microsoft 'FinancialSample.xlsx' (see https://go.microsoft.com/fwlink/?LinkID=521962)
7. Test for 1st, last & one in between piece of data
8. TODO: Run edge cases as needed
9. TODO: Enhance error handle/reporting
10. TODO: Snag a bunch of sample data and look for "unknown unknowns"