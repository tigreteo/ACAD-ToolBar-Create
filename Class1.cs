using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Customization;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolBar_Create
{
    public class Class1
    {
        public static void setUpCui()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database db = acDoc.Database;            

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //MenuGroupItemsCollection menuGcollection = tr.GetObject(db.Menu, OpenMode.ForNotify) as MenuGroupItemsCollection;

                string mainCui = Application.GetSystemVariable("MENUNAME") + ".cuix";
                string menuName = Application.GetSystemVariable("MENUNAME").ToString();
                string curWorkspace = Application.GetSystemVariable("WSCURRENT").ToString();

                CustomizationSection cs = new CustomizationSection(mainCui);
                CustomizationSection newCS = new CustomizationSection();
                PartialCuiFileCollection cuiPD = cs.PartialCuiFiles;

                WSRibbonRoot rr = cs.getWorkspace(curWorkspace).WorkspaceRibbonRoot;
                //only works with Vanilla set up menugroup = ACAD and 3D Modeling workspace
                WSRibbonTabSourceReference tab = rr.FindTabReference("ACAD", "ID_TabHome3D");

                //Toolbar toolBarPD = cs.MenuGroup.Toolbars.FindToolbarWithName("not found");

                //try to get the tool bar first else create it
                //Toolbar someToolbar = cs.MenuGroup.Toolbars.FindToolbarWithName(“tbPD”);
                //Create new ToolBar
                Toolbar tbPD = new Toolbar("New Toolbar", cs.MenuGroup);
                tbPD.ElementID = "EID_NewToolbar";
                tbPD.ToolbarOrient = ToolbarOrient.floating;
                tbPD.ToolbarVisible = ToolbarVisible.show;

                //add stuff to tool bar
                ToolbarButton newButton = new ToolbarButton(tbPD, -1);
                newButton.MacroID = "ID_Pline";
                ToolbarControl newControl = new ToolbarControl(ControlType.NamedViewControl, tbPD, -1);
                ToolbarFlyout newFlyout = new ToolbarFlyout(tbPD, -1);
                newFlyout.ToolbarReference = "Something";

                //Macro example:
                //Name:1-25in Triangular Plastic Leg
                //Macro:^C^C_quickinsert;1-25in Triangular Plastic Leg;
                //Element ID:MMU_191_1404B
                //set up new macros to add to ACAD
                MacroGroup mg = new MacroGroup("PD CUI", newCS.MenuGroup);
                MenuMacro mm1 = new MenuMacro(mg, "Cmd 1", "^C^Cmd1", "ID_MyCmd1");

                //pull down menu
                StringCollection sc = new StringCollection();
                sc.Add("needAName");
                PopMenu pm = new PopMenu("myCuiSectionName", sc, "ID_MyPop1", newCS.MenuGroup);
                PopMenuItem pmi1 = new PopMenuItem(mm1, "Pop Cmd 1", pm, -1);

            }
        }

        private void purchPartsMenuBar(string[] macroArray, CustomizationSection newCS)
        {
            MacroGroup mg = new MacroGroup("PD CUI", newCS.MenuGroup);            

            //make a macro for each name            
            MenuMacroCollection macrosCol = new MenuMacroCollection();
            foreach(string s in macroArray)
            {
                MenuMacro newMacro = new MenuMacro(mg, s, "^C^C_quickinsert;"+s+";", "ID_String");
                macrosCol.Add(newMacro);
            }

            StringCollection purchasedParts = new StringCollection();
            purchasedParts.Add("Pop15");

            //add macros to the pop menu
            PopMenu pm = new PopMenu(newCS.CUIFileName, purchasedParts, "ID_MyPop15", newCS.MenuGroup);
            int index = -1;
            foreach(MenuMacro mm in macrosCol)
            {
                PopMenuItem pmi = new PopMenuItem(mm, "Pop CMD " + index +2, pm, index);
                index++;
            }
        }


        //Older technique to import cui, can be replaced with newer AUTOLOADER feature with XML
        private void LoadMyCui(string cuiFile)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;

            object oldCmdEcho = Application.GetSystemVariable("CMDECHO");
            object oldFileDia = Application.GetSystemVariable("FILEDIA");

            Application.SetSystemVariable("CMDECHO", 0);
            Application.SetSystemVariable("FILEDIA", 0);

            doc.SendStringToExecute("_.cuiload " + cuiFile  + " ",false, false, false);
            doc.SendStringToExecute("(setvar \"FILEDIA\" "+ oldFileDia.ToString()+ ")(princ) ",false, false, false);
            doc.SendStringToExecute("(setvar \"CMDECHO\" "+ oldCmdEcho.ToString()+ ")(princ) ",false, false, false);
        }
    }
}
