using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZYCJ.Model
{
    public class AcademicAchievements
    {
        public AcademicAchievements(string professor) {
            this.professor = professor;
            this.paperInfo = new StringBuilder();
            this.paperInfo.Append("一、论文：");
            this.crosswiseProject = new StringBuilder();
            this.crosswiseProject.Append("\r\n 二、课题：");
            this.lengthwaysProject = new StringBuilder();
    }
        public string professor;
        public StringBuilder paperInfo;
        public StringBuilder crosswiseProject;
        public StringBuilder lengthwaysProject;
    }
}
