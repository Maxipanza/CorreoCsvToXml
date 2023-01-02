// See https://using Microsoft.VisualBasic.FileIO;
using System.ComponentModel.Design.Serialization;
using System.Globalization;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Net;
using System.Net.Mail;
using static CorreoCsvToXml.Utils;
using System.Reflection;
// This will get the current WORKING directory (i.e. \bin\Debug)
string workingDirectory = Environment.CurrentDirectory;
// This will get the current PROJECT directory
string projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;
var csvFileName = "tuki.csv";
var csvFile = @$"{projectDirectory}\{csvFileName}";
string rutaDescargas = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";


var parser = ReadCSVFromDirectory("tuki.csv");

parser.HasFieldsEnclosedInQuotes = true;
parser.SetDelimiters(",");

string[] fields;
XmlDocument root = new XmlDocument();
XmlElement brands = root.CreateElement("marcas");
brands.SetAttribute("xmlns", "http://chat.soybot.com/catalogo/V1");

while (!parser.EndOfData)
{
    XmlElement brand = null;
    XmlElement model = null;
    XmlElement version = null;
    XmlElement unidad = null;
    XmlElement imagenes = null;
    fields = parser.ReadFields();
    for (var i = 0; i < fields.Length; i++)
    {

        if (i == 0)
        {
            fields[i] = Capitalize(fields[i]);
            brand = AppendMarca(brands, fields[i], root);
        }
        if (i == 1 && brand != null && IsNotEmpty(fields[i]))
        {
            fields[i] = Capitalize(fields[i]);
            model = AppendModel(brand, fields[i], fields[0], root);
            brand.AppendChild(model);
        }
        if (i == 2 && model != null && IsNotEmpty(fields[i]))
        {
            fields[i] = Capitalize(fields[i]);
            version = AppendVersion(brand, fields[i], fields[0], fields[1], root);
            model.AppendChild(version);
        }
        if (i == 3 && version != null && IsNotEmpty(fields[i]))
        {
            fields[i] = Capitalize(fields[i]);
            //TODO: desacoplar el concepto de anio con unidad en este método
            unidad = AppendUnidad(fields[i], root);
            version.AppendChild(unidad);
        }
        if (i == 4 && unidad != null && IsNotEmpty(fields[i]))
        {
            //TODO: desacoplar el concepto de precio con unidad en este método
            unidad.SetAttribute("precio", fields[i]);
            //imagenes = AppendImagenes(unidad, fields[i], fields[0], fields[1], fields[2], root);
        }
        if (i == 5 && unidad != null && IsNotEmpty(fields[i]))
        {
            fields[i] = fields[i].ToLower();
            XmlElement url = root.CreateElement("url");
            url.InnerText = fields[i];
            url.SetAttribute("tipo", "foto-agencia");
            XmlElement imagen = root.CreateElement("imagenes");
            imagen.AppendChild(url);
            unidad.AppendChild(imagen);
        }
        if (i == 6 && unidad != null)
        {
            if (IsNotEmpty(fields[i]))
            {
                unidad.SetAttribute("id", fields[i]);
            }
            else
            {
                unidad.SetAttribute("id", Guid.NewGuid().ToString());
            }
        }
        if (i == 7 && unidad != null && IsNotEmpty(fields[i]))
        {
            unidad.SetAttribute("kilometros", fields[i]);
        }
        if (i == 8 && unidad != null && IsNotEmpty(fields[i]))
        {
            unidad.SetAttribute("tipoCambio", fields[i]);
        }
        if (i == 9 && unidad != null && IsNotEmpty(fields[i]))
        {
            fields[i] = Capitalize(fields[i]);
            unidad.SetAttribute("zona", fields[i]);
        }
        if (i == 10 && unidad != null && IsNotEmpty(fields[i]))
        {
            unidad.SetAttribute("lat", fields[i]);
        }
        if (i == 11 && unidad != null && IsNotEmpty(fields[i]))
        {
            unidad.SetAttribute("long", fields[i]);
        }
        if (i == 12 && unidad != null && IsNotEmpty(fields[i]))
        {

            unidad.SetAttribute("cliente", fields[i]);
        }
        if (i == 13 && unidad != null && IsNotEmpty(fields[i]))
        {
            unidad.SetAttribute("tipoVenta", fields[i]);
        }
        brands.AppendChild(brand);
    }
    root.AppendChild(brands);
}

bool IsNotEmpty(string str) => str.Length > 0;


XmlElement AppendUnidad(string anio, XmlDocument root)
{
    XmlElement unidad = root.CreateElement("unidad");
    unidad.SetAttribute("anio", anio);
    return unidad;
}

XmlElement AppendVersion(XmlElement brand, string versionName, string brandName, string modelName, XmlDocument root)
{
    if (brand.SelectSingleNode($"//marca[@nombre='{brandName}']/modelo[@display='{modelName}']/version[@display='{versionName}']") is XmlElement existingVersion)
    {
        return existingVersion;
    }
    XmlElement version = root.CreateElement("version");
    version.SetAttribute("display", versionName);
    version.SetAttribute("estado", "activo");
    version.SetAttribute("enLista", "activo");
    version.SetAttribute("id", Guid.NewGuid().ToString());
    return version;
}
XmlElement AppendModel(XmlElement brand, string name, string brandname, XmlDocument root)
{
    if (brand.SelectSingleNode($"//marca[@nombre='{brandname}']/modelo[@display='{name}']") is XmlElement existingModel)
    {
        return existingModel;
    }
    XmlElement model = root.CreateElement("modelo");
    model.SetAttribute("display", name);
    model.SetAttribute("estado", "activo");
    model.SetAttribute("enLista", "activo");
    model.SetAttribute("id", Guid.NewGuid().ToString());
    return model;
}

XmlElement AppendMarca(XmlElement brands, string name, XmlDocument root)
{
    if (brands.SelectSingleNode($"//marca[@nombre='{name}']") is XmlElement existingBrand)
    {
        return existingBrand;
    }
    XmlElement brand = root.CreateElement("marca");
    brand.SetAttribute("nombre", name);
    brand.SetAttribute("estado", "activo");
    brand.SetAttribute("enLista", "activo");
    return brand;
}

string FormatXML(string xml)
{
    try
    {
        XDocument doc = XDocument.Parse(xml);
        doc.Save("foo.xml");
        return doc.ToString();
    }
    catch (Exception e)
    {
        Console.WriteLine($"Error: Exception {e} while formatting xml");
        return xml;
    }
}

Console.WriteLine("Ingrese el nombre del archivo XML: ");
string nombreArchivo = Console.ReadLine();
Console.WriteLine(FormatXML(root.OuterXml));
var archivoGuardado = $@"{rutaDescargas}\{nombreArchivo}.xml";
root.Save(archivoGuardado);
try
{
    MailMessage mail = new MailMessage();
    SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

    mail.From = new MailAddress("runonenotification@gmail.com");
    mail.To.Add("maximiliano@run0km.com");
    mail.Subject = "envio de csv";
    mail.Body = "Generaste el csv de:" + nombreArchivo;

    Attachment attachment = new Attachment(archivoGuardado);
    mail.Attachments.Add(attachment);

    SmtpServer.Port = 587;
    SmtpServer.Credentials = new NetworkCredential("runonenotification@gmail.com", "hkkbsiazgkppinwy");
    SmtpServer.EnableSsl = true;

    SmtpServer.Send(mail);
    Console.WriteLine("Se envio el correo");
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
}
parser.Close();

/*
XElement cust = new XElement("marcas",
    from str in source
    let fields = str.Split(',')
    select new XElement("marca",
        new XAttribute("nombre", fields[0]),
        new XAttribute("estado", "activo"),
        new XElement("modelo",
            new XAttribute("display", fields[1]),
            new XAttribute("estado", "activo"),
            new XAttribute("enLista", "activo"),
            new XAttribute("grupos", fields[3]),
                new XElement("version",
                    new XAttribute("display", fields[2]),
                    new XAttribute("estado", "activo"),
                    new XAttribute("enLista", "activo"),
                new XElement("unidad",
                    new XAttribute("anio", fields[4]),
                    new XAttribute("precio", fields[5]),
                    new XAttribute("zona", fields[16]),
                new XElement("imagenes",
                    new XElement("url",
                    new XAttribute("tipo","foto-agencia"),
                    fields[9])
            )
        )

    )
*/
