using System.ComponentModel;
using System.Security.Cryptography;
using System.Text;
using OfficeOpenXml;

namespace ExcelCompareApp;

public partial class MainPage : ContentPage
{
    Stream excelStream1;
    Stream excelStream2;


    public MainPage()
    {
        InitializeComponent();
    }


    private async void PickExcel_Clicked1(object sender, EventArgs e)
    {

        var result = await FilePicker.PickAsync(new PickOptions
        {
        });

        if (result == null)
            return;

        await DisplayAlert("You picked...", result.FileName, "OK");

        excelStream1 = await result.OpenReadAsync();


    }

    private async void PickExcel_Clicked2(object sender, EventArgs e)
    {

        var result = await FilePicker.PickAsync(new PickOptions
        {
        });

        if (result == null)
            return;

        await DisplayAlert("You picked...", result.FileName, "OK");

        excelStream2 = await result.OpenReadAsync();


    }


    private void CompareExcelFiles(object sender, EventArgs e)
    {
        try
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Create a temporary file to save the stream
            string tempFilePath1 = Path.GetTempFileName();
            string tempFilePath2 = Path.GetTempFileName();

            // Write the stream data to the temporary file
            using (FileStream tempFileStream1 = File.OpenWrite(tempFilePath1))
            {
                excelStream1.CopyTo(tempFileStream1);
            }

            using (FileStream tempFileStream2 = File.OpenWrite(tempFilePath2))
            {
                excelStream2.CopyTo(tempFileStream2);
            }

            List<string> hashList1 = CalculateRowsHashes(tempFilePath1);
            List<string> hashList2 = CalculateRowsHashes(tempFilePath2);

            string comparisonResult;

            if (hashList1.Count != hashList2.Count)
            {
                comparisonResult = "Excel files have different numbers of rows.";
            }
            else if (hashList1.SequenceEqual(hashList2))
            {
                comparisonResult = "Excel files have the same rows (possibly different arrangement).";
            }
            else
            {
                comparisonResult = "Excel files have different rows or arrangements.";
            }

            ComparisonResultLabel.Text = comparisonResult;

            // Clean up: Delete the temporary file
            File.Delete(tempFilePath1);
            File.Delete(tempFilePath2);

        }
        catch (Exception ex)
        {
            ComparisonResultLabel.Text = $"Error comparing Excel files: {ex.Message}";
        }
    }

    private List<string> CalculateRowsHashes(string filePath)
    {
        List<string> hashes = new List<string>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming you're working with the first worksheet

            for (int row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                string rowHash = CalculateRowHash(worksheet, row);
                hashes.Add(rowHash);
            }
        }

        return hashes;
    }

    private string CalculateRowHash(ExcelWorksheet worksheet, int row)
    {
        var rowCells = worksheet.Cells[row, worksheet.Dimension.Start.Column, row, worksheet.Dimension.End.Column];
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(GetCellValues(rowCells)));
            return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
        }
    }

    private string GetCellValues(ExcelRangeBase cells)
    {
        StringBuilder cellValues = new StringBuilder();
        foreach (var cell in cells)
        {
            cellValues.Append(cell.Text);
        }
        return cellValues.ToString();
    }


}


