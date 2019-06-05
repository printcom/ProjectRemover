using System;

namespace ProjectRemover.Package.Classes
{
    public class RemovableProject
    {
        public Guid Id { get; set; }

        public string Name { get; set; }

        public string RelativePath { get; set; }

        public string FullPath { get; set; }

        public string NestedPath { get; set; }

        public bool Remove { get; set; }
    }
}