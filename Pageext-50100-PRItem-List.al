pageextension 50100 PRItemListExt extends "Item List"
{
    actions
    {
        addfirst(processing)
        {
            action(ImportEntityTextFromExcel)
            {
                Caption = 'Import Entity Text from Excel';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Import;
                trigger OnAction()
                begin
                    ReadExcelSheet();
                    ImportEntityTextFromExcel();
                end;
            }
            action(ExportEntityTextToExcel)
            {
                Caption = 'Export Entity Text to Excel';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Export;
                trigger OnAction()
                begin
                    ExportEntityTextToExcel();
                end;
            }
            action(ClearAllEntityText)
            {
                Caption = 'Clear All Entity Text';
                ApplicationArea = All;
                Promoted = true;
                PromotedCategory = Process;
                PromotedIsBig = true;
                Image = Delete;
                trigger OnAction()
                begin
                    ClearAllEntityText(Rec);
                end;
            }
        }
    }
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
        UploadExcelMsg: Label 'Please Choose the Excel file.';
        NoFileFoundMsg: Label 'No Excel file found!';
        ExcelImportSuccess: Label 'Excel is successfully imported.';
        ExcelExportSuccess: Label 'Entity Text data has been exported to Excel.';

    local procedure ImportEntityTextFromExcel()
    var
        RowNo: Integer;
        MaxRowNo: Integer;
        EntityText: Record "Entity Text";
        Item: Record Item;
        ScenarioValue: Enum "Entity Text Scenario";
    begin
        RowNo := 0;
        MaxRowNo := 0;
        TempExcelBuffer.Reset();
        if TempExcelBuffer.FindLast() then
            MaxRowNo := TempExcelBuffer."Row No.";

        for RowNo := 2 to MaxRowNo do begin
            if not Item.Get(GetValueAtCell(RowNo, 1)) then
                Error('Item %1 not found.', GetValueAtCell(RowNo, 1));

            if not Evaluate(ScenarioValue, GetValueAtCell(RowNo, 2)) then
                Error('Invalid Scenario value: %1', GetValueAtCell(RowNo, 2));

            EntityText.Reset();
            EntityText.SetRange(Company, CompanyName);
            EntityText.SetRange("Source Table Id", Database::Item);
            EntityText.SetRange("Source System Id", Item.SystemId);
            EntityText.SetRange(Scenario, ScenarioValue);

            if EntityText.FindFirst() then begin
                // Update existing record
                EntityText."Preview Text" := CopyStr(GetValueAtCell(RowNo, 3), 1, 1024);
                EntityText.Text.CreateOutStream(OutStr);
                OutStr.WriteText(GetValueAtCell(RowNo, 3));
                EntityText.Modify();
            end else begin
                // Insert new record
                EntityText.Init();
                EntityText.Company := CompanyName;
                EntityText."Source Table Id" := Database::Item;
                EntityText."Source System Id" := Item.SystemId;
                EntityText.Scenario := ScenarioValue;
                EntityText."Preview Text" := CopyStr(GetValueAtCell(RowNo, 3), 1, 1024);
                EntityText.Text.CreateOutStream(OutStr);
                OutStr.WriteText(GetValueAtCell(RowNo, 3));
                EntityText.Insert();
            end;
        end;
        Message(ExcelImportSuccess);
    end;

    local procedure ExportEntityTextToExcel()
    var
        EntityText: Record "Entity Text";
        Item: Record Item;
        TempBlob: Codeunit "Temp Blob";
        FileMgt: Codeunit "File Management";
        OutStream: OutStream;
        InStream: InStream;
        RowNo: Integer;
        FileName: Text;
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();

        // Add header row
        RowNo := 1;
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn('Item No.', false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn('Scenario', false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn('Text Content', false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);

        // Add data rows
        EntityText.Reset();
        EntityText.SetRange("Source Table Id", Database::Item);
        if EntityText.FindSet() then
            repeat
                if Item.GetBySystemId(EntityText."Source System Id") then begin
                    RowNo += 1;
                    TempExcelBuffer.NewRow();

                    TempExcelBuffer.AddColumn(Item."No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                    TempExcelBuffer.AddColumn(Format(EntityText.Scenario), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                    TempExcelBuffer.AddColumn(EntityText."Preview Text", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                end;
            until EntityText.Next() = 0;

        // Create Excel file
        TempExcelBuffer.CreateNewBook('Entity Text');
        TempExcelBuffer.WriteSheet('Entity Text', CompanyName, UserId);
        TempExcelBuffer.CloseBook();

        TempExcelBuffer.SetFriendlyFilename('EntityTextExport.xlsx');
        TempExcelBuffer.OpenExcel();

        Message(ExcelExportSuccess);
    end;

    local procedure ClearAllEntityText(Item: Record Item)
    var
        EntityText: Record "Entity Text";
    begin
        Item.Reset();
        if Item.FindSet() then
            repeat
                EntityText.Reset();
                EntityText.SetRange("Source Table Id", Database::Item);
                EntityText.SetRange("Source System Id", Item.SystemId);
                if EntityText.FindSet() then
                    EntityText.DeleteAll();
            until Item.Next() = 0;
    end;

    local procedure ReadExcelSheet()
    var
        FileMgt: Codeunit "File Management";
        IStream: InStream;
        FromFile: Text[100];
    begin
        UploadIntoStream(UploadExcelMsg, '', '', FromFile, IStream);
        if FromFile = '' then
            Error(NoFileFoundMsg);
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.OpenBookStream(IStream, TempExcelBuffer.SelectSheetsNameStream(IStream));
        TempExcelBuffer.ReadSheet();
    end;

    local procedure GetValueAtCell(RowNo: Integer; ColNo: Integer): Text
    begin
        TempExcelBuffer.Reset();
        If TempExcelBuffer.Get(RowNo, ColNo) then
            exit(TempExcelBuffer."Cell Value as Text")
        else
            exit('');
    end;

    var
        OutStr: OutStream;
}