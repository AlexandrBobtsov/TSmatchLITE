using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
//-- мои модули
using Decl = TSmatch.Declaration.Declaration;
using Log = match.Lib.Log;
using TS = TSmatch.Tekla.Tekla;
using FileOp = match.FileOp.FileOp;
//-- низкоуровневые обращения, от которых хорошо бы избавиться
using Excel = Microsoft.Office.Interop.Excel;
using Mtr = match.Matrix.Matr;

namespace TSmatchLITE
{
    class Program
    {
        static void Main(string[] args)
        {
            TS.GetTeklaDir();
            Elements = TS.Read();
            Ignore0Length();
            int ok = 0;
            ok += Search("Уголок", "М:С255=Ст3Гпс,Cт3гсп; П: Уголок = L $p1x$p2; lng");
            ok += Search("Полоса", "М:С255=Ст3Гпс,Cт3гсп ; П:Полоса = — $p1*$p2");
            //            ok += Search("Полоса", "М:С255=Ст3; П:Полоса = — $p1"); ..РАЗОБРАТЬСЯ с $1 и $1*$2
            ok += Search("Балка", "М:С255=Ст3Гпс=Cт3гсп; П:Балка = I Б $p1");
            // ввести 'Ш1' или "K2"  и & - обязательные элементы в разбор регулярных выражений
            ok += Search("Труба профильная", "М:С255=Ст3Гпс=Cт3гсп; П:Гнз p1xp2xp3");
            okSave(TS.ModInfo.ModelPath, ok);
            new Log("=======================\nвсего подобрали комплектующие для " + ok + " элементов");
            Console.ReadLine();
        }

        static List<TS.AttSet> Elements = new List<TS.AttSet>();

        static void Ignore0Length()
        {
            int cnt = Elements.Count;
            int cnt_initial = cnt;
            for(int i=0; i< cnt; i++)
            {
                if (Elements[i].lng <= 0)
                {
                    Elements.Remove(Elements[i]);
                    cnt--;
                }
            }
            new Log("В модели " + (cnt_initial - cnt) + " элементов нулевой длины -- игнорируем их");
        }

        static int Search(string ShtN, string rule)
        {
            int ok = 0;
            List<string> Comp = GetCompList(ShtN);
            RuleParser(rule);
            int nstr = 0;
            foreach (var v in Elements)
            {
                if (SearchInComp(ShtN, Comp, nstr, v.mat, v.prf)) ok++;
                nstr++;
            }
            new Log("в каталоге " + ShtN + "\tподобрали соответствие для " + ok + " элементов");
            return ok;
        }
        public static List<string> GetCompList(string SheetN)
        {
            const int i0 = 7;
            const string FileName = @"black1_0.xls";
            //--- Catalog Path for PavelMSI ---
            //const string FileDir = @"C:\Users\Pavel_Khrapkin\Desktop\TSmatch\База Севзапметалл";
            //--- Catalog Path for Bureau ESG office ---
            string exceldesightPath = TS.GetTeklaDir();
            string FileDir = exceldesightPath + @"\База Севзапметалл";
            List<string> Comp = new List<string>();
            Excel.Workbook Wb = FileOp.fileOpen(FileDir, FileName);
            Excel.Worksheet Sh = Wb.Worksheets[SheetN];
            Mtr Body = FileOp.getSheetValue(Sh);
            for (int i = i0; i <= Body.iEOL(); i++)
                Comp.Add(Body.Strng(i, 1));
            return Comp;
        }
        public static void okSave(string FileDir, int ok)
        {
            const string SheetN = "Report";
            const string FileName = "TSmatchINFO.xlsx";
            Excel.Workbook Wb = FileOp.fileOpen(FileDir, FileName, create_ifnotexist: true);
            FileOp.SheetReset(Wb, SheetN); 
            Excel.Worksheet Sh = Wb.Worksheets[SheetN];
//            Mtr Body = null;
            int n_ex_str = 0;
            string mat = "Материал", prf = "Профиль", lng = "Длина",
                Comp_doc = "Компонент", Comp_n = "№ стр", Comp_str = "Cтрока компонента";
            Mtr Body = new Mtr();
            Body.Init(new object[] { "№", mat, prf, lng, Comp_doc, Comp_n, Comp_str });
            foreach (var v in Elements)
            {
                mat = v.mat;
                prf = v.prf;
                lng = v.lng.ToString();
                for (int i_ok = 0; i_ok < OKnstrComp.Count; i_ok++)
                {
                    if (n_ex_str == OKnstrComp[i_ok])
                    {
                        Comp_doc = OKdocName[i_ok];
                        Comp_n = OKnComp[i_ok].ToString();
                        Comp_str = OKstrComp[i_ok];
                        break;
                    }
                    else { Comp_doc = ""; Comp_str = ""; Comp_n = ""; }
                }
                Body.AddRow(new object[] { n_ex_str, mat, prf, lng, Comp_doc, Comp_n, Comp_str });
                n_ex_str++;
            }
            FileOp.setRange(Sh);
            FileOp.saveRngValue(Body);
        }
        /// <remarks> 26.1.2016
        /// RuleParser разбирает текстовую строку - правило, выделяя и возвращая в списках 
        ///     RuleMatList - части Правила, относящиеся к Материалу компонента
        ///     RulePrfList - части, относящиеся к Профилю
        ///     RuleOthList - остальные разделы Правила.
        ///     
        ///     RuleMatMust, RulePrfMust и RuleOthMust  - обязательные элементы раздела, если есть
        ///     RuleMatNpars, RulePrfNpars и RuleOthNpars - количество числовых параметров в разделе
        ///     
        /// Разделы начинаются признаком раздела, например, "Материал:" и отделяются ';'
        /// Признак раздела распознается по первой букве 'M' и завершается ':'. Поэтому
        ///             "Профиль:" = "П:" = "Прф:" = "п:" = "Prof:"
        /// Заглавные и строчные буквы эквивалентны, национальные символы транслитерируются.
        /// Разделы Правила можно менять местами и пропускать; тогда они работают "по умолчанию".
        /// '=' означает эквивалентность. Допустимы выражения "C255=Ст3=сталь3".
        /// ',' позволяет перечислять элементы и подразделы Правила.
        /// ' ' пробелы, табы и знаки конца строки игнорируются, однако они могут служить признаком
        ///     перечисления, так же, как пробелы. Таким образом, названия материалов или профилей
        ///     можно просто перечислить - это эквивалентно '='
        /// &'..' или &".." - обязательные части при подборе соответствия. Если части, разделенные 
        ///     знаками '=', запятыми или пробелами - это альтернативы, возможные варианты, то при
        ///     указании &'..' ' наличие такой части в результате обязательно.
        ///     Обязательный элемент в разделе Правила может быть толькот один.   
        /// Параметры   - последовательность символов вида "$p1" или просто "p1" или "Параметр235".
        ///               Параметры начинаются со знака '$' или 'p'  и кончаются цифрой. 
        ///               Их значения подставляются из значений и атрибутов компонента в модели.
        ///               Номер параметра в правиле неважен- он заменяется в TSmatch на номер
        ///               по порядку следования параметров в правиле автоматически.
        /// Наличие двух и более признаков разделов без разделяющго знака ';' - ошибка.
        /// </remarks>
        static List<string> RuleMatList = new List<string>();
        static List<string> RulePrfList = new List<string>();
        static List<string> RuleOthList = new List<string>();

        static string RuleMatMust, RulePrfMust, RuleOthMust;
        static int RuleMatNpars = 0, RulePrfNpars = 0, RuleOthNpars = 0;

        static void RuleParser(string rule)
        {
            Log.set("RuleParser(\"" + rule + "\")");
            const string rM = "(m|м).*:", rP = "(п|p).*:";
            Regex regM = new Regex(rM, RegexOptions.IgnoreCase);
            Regex regP = new Regex(rP, RegexOptions.IgnoreCase);
            RuleMatList.Clear(); RulePrfList.Clear(); RuleOthList.Clear();
            RuleMatNpars = RulePrfNpars = RuleOthNpars = 0;
            string[] tmp = rule.Split(';');
            foreach (var s in tmp)
            {
                string x = s;
                while (x.Length > 0)
                {
                    if (Regex.IsMatch(x, rM, RegexOptions.IgnoreCase))
                        x = attParse(x, regM, RuleMatList, ref RuleMatNpars);
                    if (Regex.IsMatch(x, rP, RegexOptions.IgnoreCase))
                        x = attParse(x, regP, RulePrfList, ref RulePrfNpars);
                    if (x != "")
                        x = attParse(x, null, RuleOthList, ref RuleOthNpars);
                }
            }
            Log.exit();
        }
        static string attParse(string str, Regex reg, List<string> lst, ref int n)
        {
            if(reg != null) str = reg.Replace(str, "");


            string[] parametrs = Regex.Split(str, Decl.ATT_DELIM);
            foreach (var par in parametrs)
            {
                if (string.IsNullOrEmpty(par)) continue;
                bool parPars = false;
                if (Regex.IsMatch(par, Decl.ATT_PARAM)) { n++; parPars = true; }
                if (!string.IsNullOrWhiteSpace(par)
                    && !Regex.IsMatch(par, Decl.ATT_DELIM)
                    && !parPars) lst.Add(par.ToUpper());
                str = str.Replace(par, "");
            }
            if (str != "") Log.FATAL("строка \"" + str + "\" разобрана не полностью");
            return str;
        }

        static List<string> OKmat = new List<string>();
        static List<string> OKprf = new List<string>();
        static List<string> OKdocName = new List<string>();
        static List<string> OKstrComp = new List<string>();
        static List<int> OKnComp = new List<int>();
        static List<int> OKnstrComp = new List<int>();
        //!!-------------------------------------------------------
        static bool SearchInComp(string docName, List<string> Comp, int nstr, string mat, string prf)
        {
            Log.set("SearchInComp");
            bool found = false;
            mat = mat.ToUpper(); prf = prf.ToUpper();
            List<int> matPars = GetPars(mat);
            List<int> prfPars = GetPars(prf);
            int n = 0;
            foreach (var s in Comp)
            {
                n++;
                if (string.IsNullOrWhiteSpace(s)) continue;
                string str = s.ToUpper();
                if (!IContains(RuleMatList, mat)) continue;
//                if (!str.Contains(mat)) continue;
                if (!IContains(RulePrfList, prf)) continue;
                List<int> CompPars = GetPars(str);
                int i = 0;
                foreach (int x in prfPars)
                {
                    if (x != CompPars[i]) break;
                    i++;
                }
                if (i < prfPars.Count) continue;

                int iok = 0;
                foreach(var ok in OKnComp)
                {
                    if (OKnstrComp[iok] == nstr)
                        Log.Warning("Этот компонент уже в OK"
                            +"\nnstr=" + nstr + " mat = " + mat + "\tprf = " + prf + "\t" + OKdocName[iok]+"\t"+OKstrComp[iok]);
                    iok++;
                }

                found = true;
                OKmat.Add(mat);
                OKprf.Add(prf);
                OKdocName.Add(docName);
                OKstrComp.Add(str);
                OKnComp.Add(n);
                OKnstrComp.Add(nstr);
                break;
            }
            Log.exit();
            return found;
        }
        /// <summary>
        /// IContains(List<string> lst, v) возвращает true, если в списке lst есть строка, содержащаяся в v
        /// </summary>
        /// <param name="lst"></param>
        /// <param name="v"></param>
        /// <returns></returns>
        static bool IContains(List<string> lst, string v)   
        {
            bool flag = false;
            foreach(string s in lst)
                if (v.Contains(s)) { flag = true; break; }
            return flag;
        }
//!!-------------------------------------------------------------
        /// <summary>
        /// GetParse(str) разбирает строку раздела компонента, выделяя параметры.
        ///         Названия материалов, профилей и другие нечисловые подстроки игнорируются.
        /// </summary>
        /// <param name="str">входная строка раздела компонента</param>
        /// <returns>List<int>возвращаемый список найденых параметров</int></returns>
        static List<int> GetPars(string str)
        {
            const string VAL = @"\d+";
            List<int> pars = new List<int>();
            string[] pvals = Regex.Split(str, Decl.ATT_DELIM);
            foreach (var v in pvals)
            {
                if (string.IsNullOrEmpty(v)) continue;
                if (Regex.IsMatch(v, VAL))
                    pars.Add(int.Parse(Regex.Match(v, VAL).Value));
            }
            return pars;
        }
    } // end class
} // end namespace