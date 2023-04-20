using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Media3D;

namespace _23DC
{
    public class  CAP
    {
        public static List<Level> Levels=new List<Level>();
        public static List<Floor> Floors = new List<Floor>();
        public static List<Wall> Walls = new List<Wall>();
        public CAP()
        {
            Levels = new List<Level>();
            Floors = new List<Floor>();
            Walls = new List<Wall>();
        }
        public class  Level
        {
            private string levelName;
            private string levelNumbar;
            private int levelValue;
            public string LevelName { get => levelName; set => levelName = value; }
            public string LevelNumbar { get => levelNumbar; set => levelNumbar = value; }
            public int LevelValue { get => levelValue; set => levelValue = value; }
        }
        public class Floor
        {
            private string levelName;
            private string levelNumbar;
            private string levelValue;
            private List<Point3D> cPoints;
            public string LevelName { get => levelName; set => levelName = value; }
            public string LevelNumbar { get => levelNumbar; set => levelNumbar = value; }
            public string LevelValue { get => levelValue; set => levelValue = value; }
            public List<Point3D> CPoints { get => cPoints; set => cPoints = value; }
        }
        public class Wall
        {
            private string levelName;
            private string levelNumbar;
            private string levelValue;
            private List<Point3D> cPoints;
            public string LevelName { get => levelName; set => levelName = value; }
            public string LevelNumbar { get => levelNumbar; set => levelNumbar = value; }
            public string LevelValue { get => levelValue; set => levelValue = value; }
            public List<Point3D> CPoints { get => cPoints; set => cPoints = value; }
        }
    }
}
