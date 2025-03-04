tell application "Microsoft Excel"
    set currentWorkbook to workbook 1
    set lastRow to (get end (get row of range "A1"))'s row
    set lastCol to (get end (get column of range "A1"))'s column
    set workbookPath to path to currentWorkbook as string
    set folderPath to POSIX path of (container of (workbookPath as alias))
end tell

set filePath to folderPath & "data.js"

-- Create and open the file for writing
set textFile to open for access (filePath as POSIX file) with write permission
set eof of textFile to 0

-- Write the initial text
write "const jsonData = [" & linefeed to textFile

-- Loop through each row and write the data to the file
repeat with r from 2 to lastRow
    tell application "Microsoft Excel"
        set cellData to "{"
        set cellData to cellData & "\"question\": \"" & value of cell 1 of row r of currentWorkbook & "\"," & linefeed
        set cellData to cellData & "\"answer\": \"" & value of cell 2 of row r of currentWorkbook & "\"," & linefeed
        set cellData to cellData & "\"tag\": \"" & value of cell 3 of row r of currentWorkbook & "\"},"
        write cellData & linefeed to textFile
    end tell
end repeat

-- Write the final text
write "];" & linefeed to textFile

-- Close the file
close access textFile

-- Show completion message
display dialog "テキストファイルの保存が完了しました。" buttons {"OK"} default button "OK"
