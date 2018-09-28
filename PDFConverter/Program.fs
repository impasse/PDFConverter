open System.Windows.Forms
open System.Threading;
open System.Drawing
open System.IO;
open System;
open Microsoft.Office.Interop.Word;
open Microsoft.FSharp.Core;

let convert_to_pdf fullPath =
    let word = new Microsoft.Office.Interop.Word.ApplicationClass()
    try
        if not <| File.Exists(fullPath) then
            printfn "File: %A not exists" fullPath
        else
            let T = true :> obj
            let F = false :> obj
            let doc = word.Documents.Open(ref (fullPath :> obj), ReadOnly = ref T, Visible = ref F, NoEncodingDialog = ref T)
            try
                let outFileName = fullPath.Replace(".docx", ".pdf")
                doc.ExportAsFixedFormat(outFileName, WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForPrint)
            with
                e -> printfn "%A" e    
            doc.Close()
    with
        e -> printfn "%A" e
    word.Quit()
        

type Watcher () =
    let icon = new NotifyIcon(Text = "空闲中", Visible = true, Icon = new Icon(@"E:\\favicon.ico"))

    let handle (e: FileSystemEventArgs) =
        printfn "%A, %A" e.FullPath e.ChangeType
        lock Watcher (fun () -> 
        (
            icon.Text <- sprintf "处理中：%A" e.Name
            convert_to_pdf e.FullPath
            icon.Text <- "空闲中"
        ))
        

    member this.Run () =
        let cwd = Path.GetDirectoryName(Application.ExecutablePath)
        printfn "Current directory: %A" cwd
        let watcher = new FileSystemWatcher(Path = cwd, Filter = "*.docx", EnableRaisingEvents = true)
        watcher.Created.Add(handle)

[<EntryPoint>]
let main argv = 
    let watcher = Watcher()
    let thread = new Thread(watcher.Run)
    thread.Start()
    Console.ReadKey() |> ignore
    0 // return an integer exit code
