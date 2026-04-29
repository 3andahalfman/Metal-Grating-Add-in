using Autodesk.Forge.DesignAutomation.Model;
using System.Collections.Generic;

namespace Interaction
{
    /// <summary>
    /// Customizable part of Publisher class.
    /// Configured for the Metal Bar Grating Design Automation workflow.
    /// </summary>
    internal partial class Publisher
    {
        /// <summary>
        /// Constants.
        /// </summary>
        private static class Constants
        {
            private const string InventorEngine2024 = "Autodesk.Inventor+2024";
            private const string InventorEngine2025 = "Autodesk.Inventor+2025";

            public static readonly string Engine = InventorEngine2024;

            public const string Description =
                "Metal Bar Grating generator — accepts a perimeter DWG and grating " +
                "parameters JSON, produces bearing bar IPTs, cross bars, band bars, " +
                "assembly IAM, and fabrication drawing IDW as a ZIP output.";

            internal static class Bundle
            {
                public static readonly string Id = "MetalBarGrating";
                public const string Label = "prod";

                public static readonly AppBundle Definition = new AppBundle
                {
                    Engine = Engine,
                    Id = Id,
                    Description = Description
                };
            }

            internal static class Activity
            {
                public static readonly string Id = Bundle.Id;
                public const string Label = Bundle.Label;
            }

            /// <summary>
            /// DA parameter names — must match the LocalName values that
            /// SampleAutomation.Run() reads/writes on the DA server.
            /// </summary>
            internal static class Parameters
            {
                /// <summary>Perimeter boundary sketch (DWG or IPT with sketch).</summary>
                public const string PerimeterDoc = nameof(PerimeterDoc);

                /// <summary>Grating parameters JSON file.</summary>
                public const string GratingParams = nameof(GratingParams);

                /// <summary>Output ZIP containing all generated files.</summary>
                public const string OutputZip = nameof(OutputZip);
            }
        }

        /// <summary>
        /// Get command line for activity.
        /// The /i flag points to the perimeter document that Inventor opens.
        /// </summary>
        private static List<string> GetActivityCommandLine()
        {
            return new List<string>
            {
                $"$(engine.path)\\InventorCoreConsole.exe " +
                $"/al \"$(appbundles[{Constants.Activity.Id}].path)\" " +
                $"/i \"$(args[{Constants.Parameters.PerimeterDoc}].path)\""
            };
        }

        /// <summary>
        /// Get activity parameters.
        /// </summary>
        private static Dictionary<string, Parameter> GetActivityParams()
        {
            return new Dictionary<string, Parameter>
            {
                {
                    Constants.Parameters.PerimeterDoc,
                    new Parameter
                    {
                        Verb = Verb.Get,
                        Description = "Perimeter boundary file (DWG or IPT with sketch)",
                        LocalName = "perimeter.dwg"
                    }
                },
                {
                    Constants.Parameters.GratingParams,
                    new Parameter
                    {
                        Verb = Verb.Get,
                        Description = "Grating parameters JSON",
                        LocalName = "gratingParams.json"
                    }
                },
                {
                    Constants.Parameters.OutputZip,
                    new Parameter
                    {
                        Verb = Verb.Put,
                        LocalName = "output.zip",
                        Description = "ZIP containing bearing bars, cross bars, band bars, assembly, and drawing",
                        Ondemand = false,
                        Required = true
                    }
                }
            };
        }

        /// <summary>
        /// Get arguments for workitem.
        /// </summary>
        private Dictionary<string, IArgument> GetWorkItemArgs(
            string perimeterDocUrl,
            string gratingParamsUrl,
            string outputZipUrl)
        {
            var parameters = GetActivityParams();
            return new Dictionary<string, IArgument>
            {
                {
                    Constants.Parameters.PerimeterDoc,
                    new XrefTreeArgument
                    {
                        LocalName = parameters[Constants.Parameters.PerimeterDoc].LocalName,
                        Verb = Verb.Get,
                        Url = perimeterDocUrl
                    }
                },
                {
                    Constants.Parameters.GratingParams,
                    new XrefTreeArgument
                    {
                        LocalName = parameters[Constants.Parameters.GratingParams].LocalName,
                        Verb = Verb.Get,
                        Url = gratingParamsUrl
                    }
                },
                {
                    Constants.Parameters.OutputZip,
                    new XrefTreeArgument
                    {
                        LocalName = parameters[Constants.Parameters.OutputZip].LocalName,
                        Verb = Verb.Put,
                        Url = outputZipUrl
                    }
                }
            };
        }
    }
}
