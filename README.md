# Workday Studio Excel Template POC

Populate fixed cells in an existing macro-enabled Excel workbook using a custom Java step in Workday Studio.

This is a small, independent proof of concept with fictional Time Off data. It demonstrates reading an existing `.xlsm` as binary data, updating selected cells, and returning the populated workbook while retaining the sample's embedded VBA. It is a local example, not a deployed integration or a general guarantee of compatibility with vendor workbooks.

## Version and support

**v0.1.0** is the initial locally validated example. This repository is shared as a community reference with no commitment to ongoing maintenance, response times, production support, or compatibility with future Studio versions. It is not endorsed or certified by Workday. Recipients are responsible for adapting and validating it for their own environment and vendor requirements.

## Verified locally

On October 22, 2026, this project was imported into **Workday Studio 2026.24.159** and run through **Run As > AUnit Test** using a **Java 8 integration project** and the **bundled Apache POI 3.16** library.

- Native AUnit result: **Runs 1/1, Errors 0, Failures 0**.
- Three fictional records were written into `Time Off!B10:F12`; `F15` retained `SUM(F10:F12)` with a cached total of **20**.
- The output opened in Excel with the expected values. The macros were disabled during inspection.
- The embedded `xl/vbaProject.bin` was byte-for-byte identical before and after.
- An independent ZIP/XML review passed 27 checks of the sample workbook's tested features.

The native AUnit security policy was unchanged. Tenant deployment, actual vendor compatibility, Orchestration, production delivery, and VBA execution were not tested. See [VALIDATION.json](VALIDATION.json) for the recorded validation scope and source/output hashes.

## Requirements

- Workday Studio with its Workday Runtime and AUnit/JUnit libraries.
- A Java 8 execution environment for the integration project. This is separate from the Java runtime that launches the Studio application.
- The tested installation supplied POI 3.16 through the Workday Runtime. No Workday or POI JARs are distributed here.

Other Studio, Java, or POI versions need their own compilation and runtime validation. The source uses POI 3.16 APIs. [Apache no longer supports POI 4.x and earlier](https://poi.apache.org/versioning.html), which includes 3.16. This example demonstrates compatibility with the tested Studio bundle; it is not a recommendation to adopt POI 3.16 for a new application. Apache's upstream support status does not establish Workday's support for its bundled runtime. Confirm the supported Workday runtime/dependency strategy before production use or changes to bundled libraries.

## Download and import

1. On the GitHub repository page, select **Code > Download ZIP** and extract it, or clone the repository.
2. In Studio, select **File > Import... > General > Existing Projects into Workspace**.
3. Select the extracted repository's root folder, which contains `.project`.
4. Select **Studio_ExcelTemplatePOC**. Choose **Copy projects into workspace** if you want a workspace copy, then select **Finish**.
5. Allow the project to build and check **Problems** for errors. Resolve the named **Workday Runtime** and **JavaSE-1.8** execution environment if prompted.

The extracted folder name may differ, but keep the embedded project name **Studio_ExcelTemplatePOC**: the AUnit annotation references that name. If a project with that name already exists, use a separate workspace or deliberately update the name and annotation together.

## Run the example

1. Expand `src/main/java > example.timeoff`.
2. Right-click `TimeOffTemplateTest.java` and select **Run As > AUnit Test**.
3. Confirm **Runs 1/1, Errors 0, Failures 0**.
4. In **Console**, find the line beginning `Populated workbook:`. It gives the output's complete path in Studio's permitted temporary directory. The test does not write the result into the source project.
5. Open that output in Excel with macros disabled. You can also inspect the included [reference output](examples/time-off-populated-reference.xlsm) from the passing native run.

Expected data in `Time Off!B10:F12`:

| Employee ID | Name | Date | Type | Hours |
| --- | --- | --- | --- | --- |
| 000123 | Alex Sample | 2026-10-20 | PTO | 8 |
| 000456 | Jordan Example | 2026-10-21 | SICK | 4 |
| 000789 | Taylor Demo | 2026-10-22 | PTO | 8 |

`F15` should display **20.00** and retain `SUM(F10:F12)`.

## How the files work together

1. **AUnit** loads `WSAR-INF/time-off-template.xlsm` as the raw binary root message part.
2. **Start** routes to the **Populate** mediation in the assembly.
3. **PopulateTemplate** resolves the Spring bean `timeOffTemplatePopulator` and calls `TimeOffTemplatePopulator.process(byte[])`.
4. **Java** opens the existing workbook, updates the input cells, recalculates the sample's ordinary worksheet formulas, restores the original column definitions, and returns the workbook bytes.
5. **AUnit** checks input/output values, the total and formula, the hidden sheet, macro-enabled type, and identical VBA bytes, then saves the output to a temporary file.

Key files:

| File | Purpose |
| --- | --- |
| [TimeOffTemplatePopulator.java](src/main/java/example/timeoff/TimeOffTemplatePopulator.java) | Processor and three fictional input records |
| [TimeOffTemplateTest.java](src/main/java/example/timeoff/TimeOffTemplateTest.java) | Binary loading, AUnit assertions, and temporary output save |
| [assembly.xml](ws/WSAR-INF/assembly.xml) | Assembly flow and Spring bean wiring |
| [time-off-template.xlsm](ws/WSAR-INF/time-off-template.xlsm) | Prepared sample input |
| [time-off-populated-reference.xlsm](examples/time-off-populated-reference.xlsm) | Output from the passing native Studio run |
| `.project`, `.classpath`, `.settings/` | Portable Studio/Eclipse import and runtime settings |

## Adapt it to a real template

The sample has three input rows, uses Excel's 1900 date system, and accepts fictional `PTO` and `SICK` codes. `process()` supplies the records; it does not retrieve Time Off data from Workday.

To use another template, adapt the sheet name, input coordinates, record capacity, date handling, allowed codes, and matching AUnit assertions. The sample deliberately requires existing styled input cells and refuses to replace formulas in the input range. The hidden `VendorConfig` and `F15` assertions belong to this sample.

Keep the workbook binary throughout processing. An earlier text-oriented Write/static-file handoff changed the workbook bytes. The AUnit resource helper in this project preserves the original input bytes.

The column-definition snapshot/restore is intentional: in this test, POI 3.16 changed explicit `collapsed=false` column flags on load without the workaround.

`evaluateAll()` recalculates ordinary worksheet formulas in this sample; review that call against the vendor's actual formulas. **Preserving VBA does not execute VBA.** Signed VBA, ActiveX, encrypted workbooks, and complex Excel features were not validated. Test the actual vendor workbook and its acceptance requirements before production use.

A complete integration still needs its Time Off retrieval, template version management, error handling, and document delivery configured.

## License and workbook attribution

The sample code and documentation use the [Apache License 2.0](LICENSE). The bundled workbooks include third-party material described in [THIRD-PARTY-NOTICES.txt](THIRD-PARTY-NOTICES.txt), which retains the upstream license and notices.

The fictional template layout and records were created for this demonstration. Its unchanged VBA project came from Apache POI's regression fixture [`45431.xlsm` at `REL_5_5_1`](https://github.com/apache/poi/blob/REL_5_5_1/test-data/spreadsheet/45431.xlsm). These are not vendor or customer files. No Workday runtime or POI library JARs are included.

Embedded VBA SHA-256, identical in the input and reference output:

```text
4ce7af872fb4d6253972df40eb3726857258fae78d496f781b017a50416e79b5
```

## References

- [Workday: Create an Assembly Project](https://doc.workday.com/admin-guide/en-us/workday-studio/integration-design/assemblies/mcy1504015910148.html)
- [Workday: AUnit assembly testing](https://doc.workday.com/admin-guide/en-us/workday-studio/integration-monitoring-and-troubleshooting/aunit-assembly-testing/avx1516737139825.html)
- [Apache POI: Spreadsheet limitations](https://poi.apache.org/components/spreadsheet/limitations.html)
- [GitHub: Download source archives](https://docs.github.com/en/repositories/working-with-files/using-files/downloading-source-code-archives)
