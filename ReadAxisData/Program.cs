
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices.Marshalling;
using System.Xml.Linq;


//string sourcePath = @"Z:\WA4_PRICING\PRICING_FIA_MODEL\DATASETS\FIA_ERM_202412_WORK\MAIN.AXS";
string destinationPath = @"Z:\WA4_PRICING\PRICING_FIA_MODEL\DATASETS\FIA_ERM_202412_WORK\MAIN.MDB";
//File.Copy(sourcePath, destinationPath, overwrite: true);

string dbPath = destinationPath;
string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Persist Security Info=False;";

string output = @"W:\PROD_DEV\Projects\FIA_AXIS_Conversion\tables.xml";

int[] tableKeys = { 481, 482, 483, 544, 545, 549, 561, 594,595,603,604,635,636,637,639,640,641,659,660,661,844,953,954,955,957,979,984,1397,1398,4003,4006,4007,4008,4009,4010,4011,4012,4013,4014,4015,4018,4021,4023,4025,4027,4028,4029,4030,4031,4033,4036,4037,4040,4041,4042,4048,4055,4060,4064,4065};
using (OleDbConnection connection = new OleDbConnection(connectionString))
{
    try
    {
        connection.Open();
        string tableTreeLinksQuery = "SELECT\r\n    LINKS.[Parent Id],\r\n    TABLES.Name AS ParentName,\r\n    LINKS.Parent,\r\n    OBJECT.Name AS ParentType,\r\n    LINKS.[Child Id],\r\n    TABLES_1.Name AS ChildName,\r\n    LINKS.Child,\r\n    OBJECT_1.Name AS ChildType,\r\n    LINKS.Module\r\nFROM\r\n    (\r\n        (\r\n            (\r\n                LINKS\r\n                INNER JOIN [OBJECT] ON LINKS.Parent = OBJECT.Type\r\n            )\r\n            INNER JOIN [OBJECT] AS OBJECT_1 ON LINKS.Child = OBJECT_1.Type\r\n        )\r\n        LEFT JOIN TABLES ON LINKS.[Parent Id] = TABLES.Id\r\n    )\r\n    LEFT JOIN TABLES AS TABLES_1 ON LINKS.[Child Id] = TABLES_1.Id\r\nWHERE\r\n    (((LINKS.Parent) = 11))\r\n";
        string tableTreeTableLinksQuery = "SELECT\r\n    TABLELINKS.[Parent Id],\r\n    INVACCOUNT.Name AS ParentName,\r\n    OBJECT.Type,\r\n    OBJECT.Name,\r\n    TABLELINKS.[Child Id],\r\n    TABLES.Name AS ChildName,\r\n    TABLELINKS.Child,\r\n    OBJECT_1.Name,\r\n    TABLELINKS.ktype,\r\n    TABLELINKS.AsUsage,\r\n    KTYPE.Name,\r\n    KTYPE.FullName,\r\n    KTYPE.ObjectType\r\nFROM\r\n    TABLES\r\n    RIGHT JOIN (\r\n        INVACCOUNT\r\n        RIGHT JOIN (\r\n            (\r\n                (\r\n                    TABLELINKS\r\n                    INNER JOIN [OBJECT] ON TABLELINKS.Parent = OBJECT.Type\r\n                )\r\n                INNER JOIN [OBJECT] AS OBJECT_1 ON TABLELINKS.Child = OBJECT_1.Type\r\n            )\r\n            INNER JOIN KTYPE ON TABLELINKS.ktype = KTYPE.KType\r\n        ) ON INVACCOUNT.Id = TABLELINKS.[Parent Id]\r\n    ) ON TABLES.Id = TABLELINKS.[Child Id]\r\nWHERE\r\n    (\r\n        ((OBJECT.Type) = 33)\r\n        AND ((TABLELINKS.Child) = 0)\r\n        AND ((TABLELINKS.AsUsage) = 0)\r\n        AND ((KTYPE.ObjectType) = 33)\r\n    )";
        string tablesQuery = "SELECT * FROM TABLES";
        DataTable linksTable = PopulateTable(connection, tableTreeLinksQuery);
        DataTable tableLinksTable = PopulateTable(connection, tableTreeTableLinksQuery);
        DataTable tables = PopulateTable(connection, tablesQuery);
        var xml =PrintXML(tableKeys, linksTable, tableLinksTable, tables);
        using(var sw = new StreamWriter(output))
        {
            sw.WriteLine(xml);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine("Error: " + ex.Message);
    }
}


Console.WriteLine("Done.");

// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");

static DataTable PopulateTable(OleDbConnection connection, string queryString)
{
    DataTable table;
    using (OleDbCommand command = new OleDbCommand(queryString, connection))
    using (OleDbDataReader reader = command.ExecuteReader())
    {
        //while (reader.Read())
        //{
        //    for (int i = 0; i < reader.FieldCount; i++)
        //    {
        //        // Example: print first column value
        //        Console.Write(reader[i].ToString());
        //        Console.Write(",");
        //    }
        //    Console.WriteLine();
        //}


        table = new DataTable();
        table.Load(reader);
    }

    return table;
}

static void CreateNode(DataTable linksTable, DataTable tableLinksTable, TreeNode node, int key)
{
    
    var results = from row in linksTable.AsEnumerable()
                  where row.Field<int>("Parent Id") == key
                  select row;
    foreach (var row in results)
    {

        var childKey = row.Field<int>("Child Id");
        var childType = row.Field<string>("ChildType");
        var childName = row.Field<string>("ChildName");
        Console.Write("\t");
        Console.Write(childKey);
        Console.Write(" - ");
        Console.Write(childType);
        Console.Write(":");
        
        if (childType == "Table")
        {
            Console.WriteLine(childName ?? "");
            var tn = new TreeNode(childKey, childName);
            node.Children.Add(tn);
            CreateNode(linksTable, tableLinksTable, tn, childKey);
        }
        else if(childType =="Investment Account")
        {
            var parentKey = childKey;
            var iaNode = new TreeNode(parentKey, "");
            node.Children.Add(iaNode);
            var tableResults = from trow in tableLinksTable.AsEnumerable()
                               where trow.Field<int>("Parent Id")==parentKey
                               select trow;
            foreach(var trow in tableResults)
            {
                iaNode.Name = trow.Field<string>("ParentName");
                var tchildKey = trow.Field<int>("Child Id");
                var tn = new TreeNode(tchildKey, trow.Field<string>("ChildName"));
                node.Children.Add(tn);

                CreateNode(linksTable, tableLinksTable, tn, tchildKey);

            }

        }

    }
}

static XElement ConvertToXml(TreeNode node)
{
    return new XElement("Node",
        new XAttribute("ID", node.Id),
        new XAttribute("TableName", node.Name),
        node.Children.ConvertAll(ConvertToXml)
    );
}

static XElement PrintXML(int[] tableKeys, DataTable linksTable, DataTable tableLinksTable, DataTable tables)
{
    var root = new TreeNode(0, "ROOT");
    foreach (var key in tableKeys)
    {
        Console.Write(key);
        Console.Write(":");
        var results = from row in linksTable.AsEnumerable()
                      where row.Field<int>("Parent Id") == key
                      select row;

        var r = results.FirstOrDefault();
        if (r != null)
        {
            string name = r.Field<string>("ParentName");
            Console.WriteLine(name);
            var node = new TreeNode(key, name);
            root.Children.Add(node);
            CreateNode(linksTable, tableLinksTable, node, key);
        }
        else
        {
            var ts = from t in tables.AsEnumerable()
                     where t.Field<int>("Id") == key
                     select t;

            var tab = ts.First();
            string name = tab.Field<string>("Name");
            Console.WriteLine(name);
            var node = new TreeNode(key, name);
            root.Children.Add(node);

        }
    }

    XElement xml = ConvertToXml(root);
    return xml;
}

public class TreeNode
{
    public TreeNode() { }
    public TreeNode(int id,string name) { Id = id; Name = name; }
    public int Id { get; set; }
    public string Name { get; set; }
    public List<TreeNode> Children { get; set; } = new List<TreeNode>();
}