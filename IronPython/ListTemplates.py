import sys
import clr
import System
import os
import sys

lang_to_id = {
              "en" : "1033"
              }

ver_to_path = {
               "2007" : r"C:\Program Files (x86)\Microsoft Office\Office12",
               "2010" : r"C:\Program Files (x86)\Microsoft Office\Office14\Visio Content",
               "2013" : r"C:\Program Files\Microsoft Office 15\root\office15\Visio Content" }

def main() :

    visio_version = "2013"

    # Load Visio and prepare to use it
    clr.AddReference("Microsoft.Office.Interop.Visio")
    import Microsoft.Office.Interop.Visio
    IVisio = Microsoft.Office.Interop.Visio    
    flags= IVisio.VisOpenSaveArgs.visOpenRO | IVisio.VisOpenSaveArgs.visOpenDocked
    visapp = IVisio.ApplicationClass()
    docs = visapp.Documents
    assert( 0 == docs.Count)

    
    stencil_path = System.IO.Path.Combine( ver_to_path[visio_version] , lang_to_id["en"] )
    vst_files = System.IO.Directory.GetFiles(stencil_path,"*.vst*")
       
    for vst_file in vst_files:
        
        doc = docs.Open( System.IO.Path.Combine( stencil_path, vst_file) )

        for d in docs :
            tokens  = [ ("Visio"+ visio_version) , doc.Name, doc.Title, d.Name, d.Title ]
            line = "|".join(tokens)
            print line

        doc.Close()
        
    # Once done, close visio
    visapp.Quit()
    
main()
