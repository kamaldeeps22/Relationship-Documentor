using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using static ScintillaNET.Style;

namespace Relationship_Documentor
{
    [Export(typeof(IXrmToolBoxPlugin)),
        ExportMetadata("Name", "Relationship Documentor"),
        ExportMetadata("Description", "View and Document entity relationships (N:1, 1:N, N:N) with complete cascade behaviors"),
        ExportMetadata("PluginType", "Documentation"),
        ExportMetadata("SmallImageBase64", "iVBORw0KGgoAAAANSUhEUgAAACAAAAAcCAMAAAA3HE0QAAAA21BMVEVHcEwTO2kOLVERNmERNF0MJUMTPWsROGQSNl8RN2EvhMoRNV4dRWQ3ldcTPWozjdERNV0TPWs3j88yfbYve7YSOWQTPWoTOmYQMlkRNV03k9czicsRMlo2jc03ktQgWpAwg8YSOmYRNF0TO2cnca4SOWQRNF0XR3gQMFU3lNk3ktUugMUOKUs6mNsSOWUSN2IPL1QTOmcjYJQVQG01iso3jMotfcAeUYMRNmI4j80UP24SOmcSOGUTPmwTPGoTPWs2k9kSN2M0j9U4ltw5mt8wic8yjNMdVossfLoBgNS1AAAAOnRSTlMA/zHpdgn+vn+Y/GICZ+HsRfZPCRGl6rw9Vt1fNz+P4Xbbh+/+sG3+HvW1jhbJfN4ayk/KNx+eN9MtRnVBegAAAWJJREFUKM+F02eTgjAQBuDoCSog3a5nr6fXK4ggiNz//0WXbCQXLOP7iZk8ZLObCUK3spoWLkTW58oR5O92u91m47puDmeL4/u+5xlqOc/AhqzDssjiOI7azAAX/ueBU9PPARTAFQA4dYGBp2GJS7XmATB0BooC35miqwCcKQBSIAsQKtEtHq8CmYLCVVCmQKYAd5AFgjwCUMunO1Awl4dlSNHwPCKqAr+D0vzGH3QUeBB4EmoHHUFui8FXL0mSmUjnTSbFRn0E9qeEM04wIcM0Jo30ssi+5Az2gxTHh3g8yw3qk5IuoFOALBChlLw1uJYA+LRNqy/Fh0MY/vZeLwOk9DW8HgaBVlE4sBX9elqxrQEIFvcWA7gxv8p+IAKDILq3OSDK/zVNjYAoipYfDPgvz9yxzRYF+/VPCgaNzFV2WxTs110Ao0nn5K28LxcA9i1TQSu5I5y9JrvbrkBM6+bL+wNw90u6v9ZvtQAAAABJRU5ErkJggg=="),
        ExportMetadata("BigImageBase64", "iVBORw0KGgoAAAANSUhEUgAAAFAAAABGCAMAAABsQOMZAAAAtFBMVEVHcEw2lNoECAw4l9wOLlQ1ktgRNV8SN2ILIDsLIDwTOWUMJEEJHDYufLQug8kRNV4SOWM5mNsXN1URM1sPLlIMJEISK0gQMVYvebEdVosOK04IFis4kdIzgr4xfrk2j9E1i8sygr8RNFwOKEoYSXo0iMcrbaA0jM8ugcYvfr45ldguerkUP24TPWsTPGoTPmwSO2gSOWUSOmcSOGQUQG80j9UyjNIRN2Iwh846mt8jZZ4qd7VqKlYxAAAAN3RSTlMA/gP+rv7f8jke+EgkHP7N6dwHonwxC482/mQRzGBN3LFyu1f+lhLy58ruhP/////////////+aVWoBAAABDFJREFUWMPlmGlXskAYhhFRk1AIEE3L3VZ2San+//96Z5gZeGZY6/3Y3Tl96Jyucz/7qCT9RWk3v5A2Xsk1PNka/ULGULfs23ElULler+dMEVYch7n8XF6ugMp1A1Wf35hNQIKMscKY40Eo4yElrqHcylXAqwAM4xaHLiUi5HzVFnIedA3PpTSiJNjPaoBRLdATDHJEN9FnXRz6jTXJkSjoJNG1BmCHmAMuZCzLaQaWeCUgYyKDGKkueSCoMOBBg35FHwKHiT6uBtImbCiKF3gckPASd1lT5ThS2+QHVcS9U53D2FjeNutgKwaYFSZjVgqZEIfj1vVkzixVKLULYkZAwIs7ACXJsVWuc/CveQHkStIJKJlzocqoFWXRISF2A0raELY28pgoEHhlAXcGShbIYdbdiugw+hnQ9rjh44CFPdTSXYFLj5aD5bAMJMQfO8yLIrQNWwsdgbLCp9B17f8DzgyXV7CUqnLYFWhagcsXGY5exiMl8f1OQNNWwZiQpb0qTUp3hxqZZWiySGFbyPJqrHGaHeZDBPM4Ioi4sQ+dG1vRh7wMFbHAtSKy5BKQrmoANA/7UUz3tscL3oGAN8iKQm8JKArKVCw+SPhTWtx6W6oExhA406sudHatgoB/kLjKSgAWBnOgpkes3QuX9PjhHMIa78f8EEVnysP/Q4GOdbwnOib++RyL1xSGvNeEqaQlITYo0HyfIqVpOk2n31/3SRSHwkllvMASG60AhgWQEgn0cpl+HaOQePQ5j8Fw6ZT2xpm9W4FD9PfdC8ZhIEamDFmU2g284VyrWEQRiNgDffj4Ms2Rl8vH9D4LIXs84F++sbe1ys0GeRAoZ0TqEBE/vo+oON7IQC92fT9fzpyaVUkjLjmUpPVLWgARcXrvhZ5+0NCHCrNh955BCnmgvHhJU2BxgEwi4k3LMj9nzVsOOSO+QocIeflKfGMpNwIzg6x8wvriiAQ5/UpGttMCDOuAkpQRLynjISJCHq1xGzCsA0qTLQuZOBwMer2P79dFUw7jglhxAibbImDMw8R+f7NuzGE+AVU3ZXISiD2st0e5DghqElQeqadTEXMO7L3tzIYq04jd6qvnvIOaMGB/cHpqqjIZeX1V/WZ9fy6KQoG9fm87aQiZrBClpmfN3TMsSsZDqiDyQM+ufS1A4iAnPizqgCRi9bZ+BAiRd9j/LLUPD6xJIV2Qz6WQEfFOaB9SFBqybzcukvWm5LD/+Sm0DwQGevNTSaZECETE/skRR4+EHIwObQ+v9YavCQaiH9iQWQ5pxLbZ+pRbFMR+jxjEAu3DgOhCzp0Oj83FQw5kBrGK9mE59NROPEBkKfzkibRtQv1gSt002fI5pLpbA6A6tDt+PsmWz+5hIPjDens0yZdphm4dfoDLTD6etg+bzZ0g0uJjbfWbbwnNp0lJT/Jf/L70H0mbuIajDTnDAAAAAElFTkSuQmCC"),
        ExportMetadata("BackgroundColor", "Lavender"),
        ExportMetadata("PrimaryFontColor", "Black"),
        ExportMetadata("SecondaryFontColor", "Gray")]
    public class MyPlugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new MyPluginControl();
        }

        public MyPlugin()
        {
            // Uncomment if using ClosedXML or other external assemblies
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            var argName = args.Name.Substring(0, args.Name.IndexOf(","));

            List<AssemblyName> refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.Where(a => a.Name == argName).FirstOrDefault();

            if (refAssembly != null)
            {
                string dir = Path.GetDirectoryName(currAssembly.Location).ToLower();
                string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
                dir = Path.Combine(dir, folder);
                var assmbPath = Path.Combine(dir, $"{argName}.dll");

                if (File.Exists(assmbPath))
                {
                    loadAssembly = Assembly.LoadFrom(assmbPath);
                }
            }
            return loadAssembly;
        }
    }
}