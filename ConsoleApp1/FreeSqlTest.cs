using DocumentFormat.OpenXml.Spreadsheet;
using FreeSql.DataAnnotations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class FreeSqlTest
    {
        static IFreeSql fsql = new FreeSql.FreeSqlBuilder()
.UseConnectionString(FreeSql.DataType.Sqlite, @"Data Source=document.db")
.UseAutoSyncStructure(true) //automatically synchronize the entity structure to the database
.Build(); //be sure to define as singleton mode
        public void Run()
        {
            fsql.Insert(new Song { Title = "test", Url = "http://test.com", CreateTime = DateTime.Now })
                .ExecuteIdentity();
            var songs = fsql.Select<Song>();
            foreach (var item in songs.ToList())
            {
                Console.WriteLine(item.Title);
            }

            //OneToOne、ManyToOne
            fsql.Select<Tag>().Where(a => a.Parent.Parent.Name == "English").ToList();

            //OneToMany
            fsql.Select<Tag>().IncludeMany(a => a.Tags, then => then.Where(sub => sub.Name == "foo")).ToList();

            //ManyToMany
            fsql.Select<Song>()
              .IncludeMany(a => a.Tags, then => then.Where(sub => sub.Name == "foo"))
              .Where(s => s.Tags.Any(t => t.Name == "Chinese"))
              .ToList();
        }

        class Song
        {
            [Column(IsIdentity = true)]
            public int Id { get; set; }
            public string Title { get; set; }
            public string Url { get; set; }
            public DateTime CreateTime { get; set; }

            public ICollection<Tag> Tags { get; set; }
        }
        class Song_tag
        {
            public int Song_id { get; set; }
            public Song Song { get; set; }

            public int Tag_id { get; set; }
            public Tag Tag { get; set; }
        }
        class Tag
        {
            [Column(IsIdentity = true)]
            public int Id { get; set; }
            public string Name { get; set; }

            public int? Parent_id { get; set; }
            public Tag Parent { get; set; }

            public ICollection<Song> Songs { get; set; }
            public ICollection<Tag> Tags { get; set; }
        }
    }
}
