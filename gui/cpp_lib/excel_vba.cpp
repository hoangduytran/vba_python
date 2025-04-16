// excel_vba.cpp
// NOTE: This code uses COM automation. It must be compiled on Windows with Microsoft Excel installed.
// Also, ensure "Trust access to the VBA project object model" is enabled in Excel.

#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <vector>
#include <string>
#include <stdexcept>
#include <comdef.h>

namespace py = pybind11;

// The following #import directives import type libraries for Excel and the VBA Extensibility objects.
// Adjust the paths if necessary depending on your Office version and installation path.
#import "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE" \
    rename("DialogBox", "ExcelDialogBox") \
    rename("RGB", "ExcelRGB")
#import "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.olb" \
    rename("EOF", "VBEOF")

// run_vba_on_all_files:
//    vba_script_file : Path to a .bas file containing VBA code.
//    excel_file_list : A list of Excel file paths.
// For each Excel file, the function:
//   - Opens the workbook.
//   - Imports the VBA module from the .bas file into the workbook (via its VBProject).
//   - Iterates over every worksheet and calls the VBA macro "RunVBA" with the worksheet name as argument.
//   - Collects the returned string from the macro.
std::vector<std::string> run_vba_on_all_files(const std::string &vba_script_file, const std::vector<std::string> &excel_file_list) {
    std::vector<std::string> results;

    // Initialize COM library
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        throw std::runtime_error("Failed to initialize COM library.");
    }

    try {
        // Create an instance of Excel
        Excel::_ApplicationPtr pXL;
        hr = pXL.CreateInstance("Excel.Application");
        if (FAILED(hr)) {
            throw std::runtime_error("Failed to create Excel application instance.");
        }

        // (Optional) Make Excel visible for debugging purposes.
        pXL->Visible = VARIANT_TRUE;

        // Loop over the provided Excel file list
        for (const auto &excelFile : excel_file_list) {
            // Open the workbook
            Excel::_WorkbookPtr pWB = pXL->Workbooks->Open(_bstr_t(excelFile.c_str()));

            // Import the VBA module from the specified .bas file.
            // Note: For this to work, programmatic access to the VBA project must be enabled.
            VBE6::VBProjectPtr pVBProj = pWB->VBProject;
            VBE6::VBComponentsPtr pVBComps = pVBProj->VBComponents;
            VBE6::VBComponentPtr pModule = pVBComps->Import(_bstr_t(vba_script_file.c_str()));

            // Get the collection of worksheets in the workbook.
            Excel::SheetsPtr pSheets = pWB->Worksheets;
            long sheetCount = 0;
            pSheets->get_Count(&sheetCount);

            // Process each worksheet. Here we assume that the imported .bas file
            // defines a function "RunVBA" that accepts a worksheet name (or other parameter)
            // and returns a value convertible to a string.
            for (long i = 1; i <= sheetCount; i++) {
                Excel::_WorksheetPtr pSheet = pSheets->Item[i];
                _variant_t vtSheetName = _variant_t(pSheet->Name);
                _variant_t vtResult = pXL->Run(_bstr_t("RunVBA"), vtSheetName);

                // Convert the result to a string and store it.
                _bstr_t bstrResult(vtResult);
                results.push_back(std::string((const char*)bstrResult));
            }

            // Optionally remove the imported module from the VBComponents.
            if (pModule) {
                pVBComps->Remove(pModule);
            }
            // Close the workbook without saving changes.
            pWB->Close(VARIANT_FALSE);
        }

        // Quit Excel.
        pXL->Quit();
    } catch (const _com_error &e) {
        CoUninitialize();
        throw std::runtime_error(std::string("COM error: ") + (const char*)e.ErrorMessage());
    }

    // Uninitialize COM library.
    CoUninitialize();
    return results;
}

// Expose the function as a pybind11 module.
PYBIND11_MODULE(excel_vba, m) {
    m.doc() = "Module to run a VBA script on Excel files via COM automation";
    m.def("run_vba_on_all_files", &run_vba_on_all_files,
          "Execute a .bas VBA script on all worksheets in provided Excel files",
          py::arg("vba_script_file"), py::arg("excel_file_list"));
}
