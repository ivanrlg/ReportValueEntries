report 50100 "Value Entry"
{
    ApplicationArea = All;
    Caption = 'Value Entry';
    UsageCategory = ReportsAndAnalysis;
    ProcessingOnly = true;
    dataset
    {
    }

    requestpage
    {
        layout
        {
            area(content)
            {
                group(GroupName)
                {
                }
            }
        }
        actions
        {
            area(processing)
            {
            }
        }
    }

    trigger OnPostReport()
    begin
        CreateExcelbook;
    end;

    local procedure RunProcess()
    var
        ValueEntry: Record "Value Entry";
        TotalByCategory, TotalByCategoryTemp : decimal;
        ItemCategory: Text;
        Item: Record Item;
        CalculatingLinesMsg: Label 'Counting Items Categories...\\';
        CurrentItemMsg: Label 'Current Entry No #1########   Current Item Category #2############', Comment = '%1,%2 = counters';
        Window: Dialog;
    begin
        ValueEntry.Reset();
        ValueEntry.SetCurrentKey("Item No.");
        ValueEntry.SetRange("Adjustment", false);
        ValueEntry.Ascending(true);
        // ValueEntry.SetRange("Posting Date", 20230101D, 20231231D);
        if ValueEntry.FindSet() then begin

            Window.Open(CalculatingLinesMsg + CurrentItemMsg);

            repeat
                Item.Get(ValueEntry."Item No.");

                ItemCategory := Item."Item Category Code";

                Window.UPDATE(1, ValueEntry."Entry No.");
                if DictionaryEntryNoByCategory.ContainsKey(ItemCategory) then begin
                    TotalByCategory += 1;
                    DictionaryEntryNoByCategory.Get(ItemCategory, TotalByCategoryTemp);
                    DictionaryEntryNoByCategory.Set(ItemCategory, TotalByCategory + TotalByCategoryTemp);
                    Window.UPDATE(2, ItemCategory);
                end else begin
                    TotalByCategory := 1;
                    DictionaryEntryNoByCategory.Add(ItemCategory, TotalByCategory);
                end;

            until ValueEntry.Next() = 0;

            Window.Close();
        end;

    end;

    local procedure CreateExcelbook()
    begin
        MakeExcelDataHeader;
        RunProcess();
        CreateExcelBody();

        ExcelBuf.CreateNewBook(Text002);
        ExcelBuf.WriteSheet(Text001, COMPANYNAME(), USERID());
        ExcelBuf.CloseBook();
        ExcelBuf.OpenExcel();
        Error('');
    end;

    local procedure MakeExcelDataHeader()
    begin
        ExcelBuf.NewRow;
        ExcelBuf.AddColumn('Total', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Item Category', FALSE, '', TRUE, FALSE, TRUE, '', ExcelBuf."Cell Type"::Text);
    end;

    local procedure CreateExcelBody()
    var
        ItemCategory: Text;
        Total: decimal;
    begin
        foreach ItemCategory in DictionaryEntryNoByCategory.Keys do begin
            DictionaryEntryNoByCategory.Get(ItemCategory, Total);

            ExcelBuf.NewRow;
            ExcelBuf.AddColumn(ItemCategory, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
            ExcelBuf.AddColumn(Total, FALSE, '', FALSE, FALSE, FALSE, '#,##0.00', ExcelBuf."Cell Type"::Number);
        end;
    end;

    var
        DictionaryEntryNoByCategory: Dictionary of [Text, decimal];
        ExcelBuf: Record "Excel Buffer" temporary;
        Text001: Label 'Data';
        Text002: Label 'Value Entry By Category';

}
