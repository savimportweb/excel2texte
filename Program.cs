using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using ExcelDataReader;
using System.IO;
using System.Data;

namespace Spreadsheetlitetest
{
    class Program
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {

            string cheminFichier = "";
            string nomFeuille = "";
            string fichierDeSortie = "";
            string strNbFeuille = "";
            //int numerofeuille = 0;
            string axeXY = "";
            bool outputFlag = false;
            //pas d'arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Vous devez préciser des arguments !\r\n\t" +
                    "-f\"Chemin Du fichier XLS/XLSX\"\r\n\t-s\"Nom de la Feuille\"\r\n" +
                    "PAS ENCORE IMPLEMENTE :-dNuméro de la feuille\r\n\t-xy\"Position X;Y de départ\", example : A1={0;0}\r\n" +
                    "t -o output dans le meme dossier avec meme nom mais .csv\r\n\t-v Version" +
                    "");
                //Console.ReadLine();
                System.Environment.Exit(-1);
            }
            else
            {

                foreach (string valeur in args)
                {
                    //Console.WriteLine(valeur + "\r\n");
                    // si - 

                    if (valeur.Contains("-v"))
                    {
                       Console.WriteLine("Version v0.10");
                       //   return (-1);
                       System.Environment.Exit(-1);
                        
                    }
                    if (valeur.Contains("-f"))
                    {
                        //cheminFichier = My_strextract(valeur, "-f", " ");
                        cheminFichier = valeur.Replace("-f", "");

                        Console.WriteLine("chemin du fichier : " + cheminFichier.ToString());
                        if (cheminFichier.Length == 0)
                        {
                            Console.WriteLine("chemin de fichier Invalide\r\\");
                            //   return (-1);
                            System.Environment.Exit(-1);
                        }
                    }
                    if (valeur.Contains("-s"))
                    {
                        //nomFeuille = My_strextract(valeur, "-s\"", "\"");
                        nomFeuille = valeur.Replace("-s", "");
                        Console.WriteLine("nomFeuille : " + nomFeuille.ToString());

                        if (nomFeuille.Length == 0)
                        {
                            Console.WriteLine("nom de feuille invalide\\n");
                            System.Environment.Exit(-1);

                        }
                    }
                    if (valeur.Contains("-o"))
                    {
                        outputFlag = true;
                    }
                    if (valeur.Contains("-d"))
                    {
                        strNbFeuille = valeur.Replace("-d", ""); ;
                    }
                    if (valeur.Contains("-v"))
                    {
                        Console.WriteLine("\tVersion_v_0_04");
                        System.Environment.Exit(-1);
                    }

                    if (valeur.Contains("-xy"))
                    {
                        //nomFeuille = my_strextract(valeur, "-s", " ");
                        axeXY = valeur.Replace("-xy", "");
                        //axeXY = My_strextract(valeur, "\"", "\"");

                        Console.WriteLine("FichierSortie : " + axeXY.ToString());

                        if (axeXY.Length == 0)
                        {
                            Console.WriteLine("axeXY invalide\\n");
                            System.Environment.Exit(-1);
                        }
                        else
                        {
                            axeXY = valeur.Replace("-xy", "");
                        }
                    }

         
                }
 
            }

            //faire gestion si pas de fichier
            IExcelDataReader excelReader;
          

            if (System.IO.File.Exists(cheminFichier) == false)
            {
                Console.WriteLine("le fichier d'entree spécifié n'existe pas\r\n");
                System.Environment.Exit(-1);

            }
            try
            {

                FileStream stream = File.Open(cheminFichier, FileMode.Open, FileAccess.Read);
                //debug pour fichier pourri ?
                
                if (Path.GetExtension(cheminFichier).ToUpper() == ".XLS")
                {
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)

                    try
                    {
                        excelReader = ExcelReaderFactory.CreateBinaryReader(stream);


                        // OK creation du dataset
                        DataSet resultat = excelReader.AsDataSet();
                        int axeX;
                        int axeY;

                        //string axeX = axeXY.Substring(1, axeXY.IndexOf(";"));
                        if (string.IsNullOrEmpty(axeXY))
                        {
                            axeX = 0;
                            axeY = 0;
                        }
                        else
                        {
                            Console.WriteLine("axeXY : " + axeXY);

                            axeX = Int32.Parse(axeXY.Substring(0, axeXY.IndexOf(";")));
                            axeY = Int32.Parse(axeXY.Substring(axeXY.IndexOf(";") + 1, axeXY.Length - axeXY.IndexOf(";") - 1));
                            Console.WriteLine("axeX : {" + axeX.ToString() + "}, axe Y : {" + axeY.ToString() + "}");
                        }

                        //recuperer le nom de la feuille 
                        //string zed = null;
                        try
                        {
                            string retourfichier = "";
                            DataTable feuille = new DataTable() ;

                            //si pas de nom de feuille on prends la premiere
                            if (string.IsNullOrEmpty(nomFeuille))
                            {
                                if (strNbFeuille.Length != 0)
                                {

                                    int nbFeuille = Int32.Parse(strNbFeuille);
                                    feuille = resultat.Tables[nbFeuille];
                                }
                                else
                                {
                                    feuille = resultat.Tables[0];
                                }
                            }
                            else
                            {
                                feuille = resultat.Tables[nomFeuille];
                            }

                            for (var countx = axeX; countx < feuille.Rows.Count; countx++)
                            {
                                for (var county = axeY; county < feuille.Columns.Count; county++)
                                {

                                    var cellule = feuille.Rows[countx][county];
                                    Console.Write(cellule.ToString() + "\t");
                                    retourfichier = retourfichier + cellule.ToString() + "\t";
                                }
                                Console.WriteLine("\r\n");
                                retourfichier = retourfichier + "\r\n";
                            }
                            if (outputFlag == false)
                            {
                                fichierDeSortie = AppDomain.CurrentDomain.BaseDirectory + "output.txt";
                            }
                            else
                            {
                                fichierDeSortie = cheminFichier.Substring(0, cheminFichier.LastIndexOf('.')) + ".csv";
                            }

                            Console.WriteLine("fichierDeSortie : " + fichierDeSortie.ToString());
                            //Console.ReadLine();
                            System.IO.File.WriteAllText(fichierDeSortie, retourfichier);

                            excelReader.Close();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("la feuille demandée n'existe pas. " + ex.ToString());
                            //Console.ReadLine();
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Signature du Fichier Invalide : " + ex.Message.ToString());
                        if (outputFlag == false)
                        {
                            fichierDeSortie = AppDomain.CurrentDomain.BaseDirectory + "output.txt";
                        }
                        else
                        {
                            fichierDeSortie = cheminFichier.Substring(0, cheminFichier.LastIndexOf('.')) + ".csv";
                        }
                        string retourfichier = "";
                        System.IO.File.WriteAllText(fichierDeSortie, retourfichier);
                        System.Environment.Exit(-1);

                    }
                }
                else
                    {
                        //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                        excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        // OK creation du dataset
                        DataSet resultat = excelReader.AsDataSet();
                        int axeX;
                        int axeY;

                        //string axeX = axeXY.Substring(1, axeXY.IndexOf(";"));
                        if (string.IsNullOrEmpty(axeXY))
                        {
                            axeX = 0;
                            axeY = 0;
                        }
                        else
                        {
                            Console.WriteLine("axeXY : " + axeXY);

                            axeX = Int32.Parse(axeXY.Substring(0, axeXY.IndexOf(";")));
                            axeY = Int32.Parse(axeXY.Substring(axeXY.IndexOf(";") + 1, axeXY.Length - axeXY.IndexOf(";") - 1));
                            Console.WriteLine("axeX : {" + axeX.ToString() + "}, axe Y : {" + axeY.ToString() + "}");
                        }

                        //recuperer le nom de la feuille 
                        //string zed = null;
                        try
                        {
                            
                            string retourfichier = "";
                            DataTable feuille = new DataTable();
                            //si pas de nom de feuille on prends la premiere
                            if (string.IsNullOrEmpty(nomFeuille))
                                {
                                    if (strNbFeuille.Length != 0)
                                    {

                                        int nbFeuille = Int32.Parse(strNbFeuille);
                                        feuille = resultat.Tables[nbFeuille];
                                    }
                                    else
                                    {
                                        feuille = resultat.Tables[0];
                                    }
                                }
                                else
                                {
                                    feuille = resultat.Tables[nomFeuille];
                                }

                                //var feuille = resultat.Tables[nomFeuille];
                                for (var countx = axeX; countx < feuille.Rows.Count; countx++)
                                    {
                                        for (var county = axeY; county < feuille.Columns.Count; county++)
                                        {

                                            var cellule = feuille.Rows[countx][county];
                                            Console.Write(cellule.ToString() + "\t");
                                            retourfichier = retourfichier + cellule.ToString() + "\t";
                                        }
                                        Console.WriteLine("\r\n");
                                        retourfichier = retourfichier + "\r\n";
                                    }
                                    if (outputFlag == false)
                                    {
                                        fichierDeSortie = AppDomain.CurrentDomain.BaseDirectory + "output.txt";
                                    }
                                    else
                                    {
                                        fichierDeSortie = cheminFichier.Substring(0, cheminFichier.LastIndexOf('.')) + ".csv";
                                    }

                                    Console.WriteLine("fichierDeSortie : " + fichierDeSortie.ToString());
                                    
                                    System.IO.File.WriteAllText(fichierDeSortie, retourfichier);

                                    excelReader.Close();
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Signature du Fichier Invalide : " + ex.Message.ToString() );
                                    if (outputFlag == false)
                                    {
                                        fichierDeSortie = AppDomain.CurrentDomain.BaseDirectory + "output.txt";
                                    }
                                    else
                                    {
                                        fichierDeSortie = cheminFichier.Substring(0, cheminFichier.LastIndexOf('.')) + ".csv";
                                    }
                                string retourfichier = "";
                                System.IO.File.WriteAllText(fichierDeSortie, retourfichier);
                                System.Environment.Exit(-1);

                                }
                    }
                }

            catch (Exception Ex)
            {
                Console.WriteLine("Exception généréé lors de l'ouverture du fichier :\r\n" + Ex.Message.ToString());
                System.Environment.Exit(-1);
            }
        }

    }


}

