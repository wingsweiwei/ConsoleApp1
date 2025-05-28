using GeoAPI.CoordinateSystems;
using GeoAPI.Geometries;
using ProjNet.CoordinateSystems;
using ProjNet.CoordinateSystems.Transformations;
using System.Globalization;
using System.Security.Principal;

namespace ConsoleApp1;

internal class GPSTest
{

    private const string HK1980_WKT = """
        GEOGCS["Hong Kong 1980",
        DATUM["Hong_Kong_1980",
            SPHEROID["International 1924",6378388,297],
            TOWGS84[-162.619,-276.959,-161.764,-0.067753,2.243648,1.158828,-1.094246]],
        PRIMEM["Greenwich",0,
            AUTHORITY["EPSG","8901"]],
        UNIT["degree",0.0174532925199433,
            AUTHORITY["EPSG","9122"]],
        AUTHORITY["EPSG","4611"]]
        """;
    private const string HK1980_Grid_WKT = """
            PROJCS["Hong Kong 1980 Grid System",
        GEOGCS["Hong Kong 1980",
            DATUM["Hong_Kong_1980",
                SPHEROID["International 1924",6378388,297],
                TOWGS84[-162.619,-276.959,-161.764,-0.067753,2.243648,1.158828,-1.094246]],
            PRIMEM["Greenwich",0,
                AUTHORITY["EPSG","8901"]],
            UNIT["degree",0.0174532925199433,
                AUTHORITY["EPSG","9122"]],
            AUTHORITY["EPSG","4611"]],
        PROJECTION["Transverse_Mercator"],
        PARAMETER["latitude_of_origin",22.3121333333333],
        PARAMETER["central_meridian",114.178555555556],
        PARAMETER["scale_factor",1],
        PARAMETER["false_easting",836694.05],
        PARAMETER["false_northing",819069.8],
        UNIT["metre",1,
            AUTHORITY["EPSG","9001"]],
        AUTHORITY["EPSG","2326"]]
        """;

    public void Run()
    {
        // 输入的经纬度字符串
        double[][][] locations =
        [
            [
                [22, 16.1476, 0], [114, 7.5529, 0]
            ],
            [
                [22, 16.1476, 0], [114, 7.5529, 0]
            ],
            [
                [22, 16.1673, 0], [114, 7.5468, 0]
            ],
            [
                [22, 16.1673, 0], [114, 7.5468, 0]
            ],
            [
                [22, 16.1677, 0], [114, 7.5257, 0]
            ],
            [
                [22, 16.1537, 0], [114, 7.5433, 0]
            ],
            [
                [22, 16.1819, 0], [114, 7.5295, 0]
            ],
        ];
        for (int i = 0; i < locations.Length; i++)
        {
            double[][]? item = locations[i];
            int num = i + 1;
            Console.WriteLine($"{num}.");
            Console.WriteLine($"原始经纬度: {item[0][0]}°{item[0][1]:00.000000}'{item[0][2]}\", {item[1][0]}°{item[1][1]:00.000000}'{item[1][2]}\"");
            // 转换为十进制度
            double latitudeDecimal = DMSToDecimal(item[0][0], item[0][1], item[0][2]);
            double longitudeDecimal = DMSToDecimal(item[1][0], item[1][1], item[1][2]);
            //Console.WriteLine($"十进制度经纬度 (十进制度): {latitudeDecimal}, {longitudeDecimal}");
            var (X, Y) = WGSToHK1980(longitudeDecimal, latitudeDecimal);
            Console.WriteLine($"转换后的XY坐标: {X}, {Y}");
            Console.WriteLine($"Json:");
            Console.WriteLine($"\"Longitude\":{longitudeDecimal},\r\n\"Latitude\":{latitudeDecimal}");
            Console.WriteLine();
        }
    }
    public (double X, double Y) WGSToHK1980(double longitude, double latitude)
    {
        var csFactory = new CoordinateSystemFactory();
        // 定义源坐标系（WGS84）
        var sourceCoordinateSystem = GeographicCoordinateSystem.WGS84;
        //var sourceCoordinateSystem = csFactory.CreateFromWkt(HK1980_WKT);
        // 定义目标坐标系（HK1980）
        var targetCoordinateSystem = csFactory.CreateFromWkt(HK1980_Grid_WKT);
        // 创建一个坐标转换对象
        CoordinateTransformationFactory ctFactory = new CoordinateTransformationFactory();
        var transform = ctFactory.CreateFromCoordinateSystems(sourceCoordinateSystem, targetCoordinateSystem);
        // 进行坐标转换
        double[] sourcePoints = [longitude, latitude];
        double[] targetPoints = transform.MathTransform.Transform(sourcePoints);

        // 输出结果
        return (targetPoints[0], targetPoints[1]);
    }
    public static double DMSToDecimal(double degree, double minute, double second)
    {
        // 检查输入是否有效
        if (minute < 0 || minute >= 60 || second < 0 || second >= 60)
        {
            throw new ArgumentException("分和秒必须在0到60的范围内");
        }

        // 将分和秒转换为度数
        double decimalValue = degree + (minute / 60.0) + (second / 3600.0);

        return decimalValue;
    }

    public (double latitude, double longitude) ParseGGA(string ggaSentence)
    {
        // 去除句子开头的$和结尾的*以及校验和
        ggaSentence = ggaSentence[1..ggaSentence.IndexOf('*', 1)];

        // 使用逗号分割句子
        string[] parts = ggaSentence.Split(',');

        // 检查句子是否为GGA类型
        if (parts[0] != "GPGGA")
            throw new ArgumentException("Invalid NMEA sentence type");

        // 解析纬度
        double latitude = ConvertDMStoDecimal(parts[2]);
        char latitudeDirection = parts[3][0];
        if (latitudeDirection == 'S')
            latitude = -latitude;

        // 解析经度
        double longitude = ConvertDMStoDecimal(parts[4]);
        char longitudeDirection = parts[5][0];
        if (longitudeDirection == 'W')
            longitude = -longitude;

        return (latitude, longitude);
    }

    private static double ConvertDMStoDecimal(string dms)
    {
        // 将DDMM.MMMM格式转换为十进制格式
        double degrees = double.Parse(dms[..^7]);
        double minutes = double.Parse(dms[^7..]);
        return degrees + (minutes / 60.0);
    }
}
