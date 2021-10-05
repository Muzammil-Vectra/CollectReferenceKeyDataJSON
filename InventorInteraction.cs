using System;
using System.Threading;
using System.Windows;
using System.Xml;
using Inventor;
using Newtonsoft.Json;
using Environment = System.Environment;
using File = System.IO.File;
using Formatting = Newtonsoft.Json.Formatting;

namespace CollectReferenceKeyDataJSON
{
    public class InventorInteraction
    {
        public static Inventor.Application InventorApp ;
        public static Document ActiveDocument;
        private readonly bool _DoNotProceed;
        public InventorInteraction()
        {
            try
            {
                InventorApp = (Inventor.Application)MarshalForCore.GetActiveObject("Inventor.Application");
                ActiveDocument = InventorApp.ActiveDocument;
            }
            catch (Exception e)
            {
                MessageBox.Show("Unable to connect. Please open Inventor and load the Document", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                _DoNotProceed = true;
            }
        }


        public static string Path = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\";
        public void CollectReferenceKeyData()
        {
            if (_DoNotProceed) return;
            if (IfFileExists())
            {
                MessageBox.Show("The File already exists", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            int edgeCounter = 0;
            int faceCounter = 0;
            int surfaceBodyCounter = 0;
            TopologyMap topologyMap = new TopologyMap();
            Double counter = 0;
            //For Assembly document
            if (ActiveDocument is AssemblyDocument)
            {
                ReferenceKeyManager referenceKeyManager = ActiveDocument.ReferenceKeyManager;
                int keyContextPointer = referenceKeyManager.CreateKeyContext();
                AssemblyDocument theAssemDoc = ActiveDocument as AssemblyDocument;
                ComponentOccurrences assemComponentOcc = theAssemDoc.ComponentDefinition.Occurrences;
               
                foreach (ComponentOccurrence componentOccurrence in assemComponentOcc)
                {
                    counter += 1;
                    MainWindow.Main.UpdateProgressBar = (counter / assemComponentOcc.Count)*100;
                    MainWindow.Main.UpdateProgressBarStatus = Math.Round(((counter / assemComponentOcc.Count) * 100),1).ToString() + "%";
                    if (componentOccurrence.SurfaceBodies.Count == 0)
                    {
                        foreach (ComponentOccurrence subComponentOccurrence in componentOccurrence.SubOccurrences)
                        {
                            foreach (SurfaceBody surfBody in subComponentOccurrence.SurfaceBodies)
                            {
                                byte[] referenceKey = new byte[] { };
                                surfBody.GetReferenceKey(ref referenceKey, keyContextPointer);
                                string sbkey = referenceKeyManager.KeyToString(ref referenceKey);
                                surfaceBodyCounter++;
                                topologyMap.BodyIdentifiers["Body_" + surfaceBodyCounter] = sbkey;

                                foreach (Face face in surfBody.Faces)
                                {
                                    face.GetReferenceKey(ref referenceKey, keyContextPointer);
                                    string faceKey = referenceKeyManager.KeyToString(ref referenceKey);
                                    faceCounter++;
                                    topologyMap.FaceIdentifiers["Face_" + faceCounter] = faceKey;
                                    foreach (Edge edge in face.Edges)
                                    {
                                        if (edge.CurveType != CurveTypeEnum.kUnknownCurve)
                                        {
                                            edge.GetReferenceKey(ref referenceKey, keyContextPointer);
                                            string edgeKey = referenceKeyManager.KeyToString(ref referenceKey);
                                            edgeCounter++;
                                            topologyMap.EdgeIdentifiers["Edge_" + edgeCounter] = edgeKey;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    foreach (SurfaceBody surfBody in componentOccurrence.SurfaceBodies)
                    {
                        try
                        {
                            byte[] referenceKey = new byte[] { };
                            surfBody.GetReferenceKey(ref referenceKey, keyContextPointer);
                            string sbkey = referenceKeyManager.KeyToString(ref referenceKey);
                            surfaceBodyCounter++;
                            topologyMap.BodyIdentifiers["Body_" + surfaceBodyCounter] = sbkey;

                            foreach (Face face in surfBody.Faces)
                            {
                                face.GetReferenceKey(ref referenceKey, keyContextPointer);
                                string faceKey = referenceKeyManager.KeyToString(ref referenceKey);
                                faceCounter++;
                                topologyMap.FaceIdentifiers["Face_" + faceCounter] = faceKey;
                                foreach (Edge edge in face.Edges)
                                {
                                    if (edge.CurveType != CurveTypeEnum.kUnknownCurve)
                                    {
                                        edge.GetReferenceKey(ref referenceKey, keyContextPointer);
                                        string edgeKey = referenceKeyManager.KeyToString(ref referenceKey);
                                        edgeCounter++;
                                        topologyMap.EdgeIdentifiers["Edge_" + edgeCounter] = edgeKey;
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            // ignored
                        }
                    }
                }

                //Save key context
                byte[] keyContextArray = new byte[] { };
                referenceKeyManager.SaveContextToArray(keyContextPointer, ref keyContextArray);
                string keyContext = referenceKeyManager.KeyToString(ref keyContextArray);
                topologyMap.AssemblyIdentifier = keyContext;
                topologyMap.AssemblyNumber = theAssemDoc.DisplayName.Split('.')[0];
                JsonSerializerSettings settings = new JsonSerializerSettings
                {
                    PreserveReferencesHandling = PreserveReferencesHandling.Objects,
                    NullValueHandling = NullValueHandling.Ignore
                };
                var json = JsonConvert.SerializeObject(topologyMap, Formatting.Indented, settings);
                File.AppendAllText(Path + topologyMap.AssemblyNumber + ".json", json);
            }



            //For part document 
            if (ActiveDocument is PartDocument)
            {
                ReferenceKeyManager referenceKeyManager = ActiveDocument.ReferenceKeyManager;
                int keyContextPointer = referenceKeyManager.CreateKeyContext();
                PartDocument ThePartDoc = ActiveDocument as PartDocument;
                PartComponentDefinition partComponentDefinition = ThePartDoc.ComponentDefinition;
                SurfaceBodies surfaceBodies = partComponentDefinition.SurfaceBodies;
                foreach (SurfaceBody body in surfaceBodies)
                {
                    counter += 1;
                    MainWindow.Main.UpdateProgressBar = (counter / surfaceBodies.Count) * 100;
                    MainWindow.Main.UpdateProgressBarStatus = Math.Round(((counter / surfaceBodies.Count) * 100), 1).ToString()+"%";
                    byte[] referenceKey = new byte[] { };
                    body.GetReferenceKey(ref referenceKey, keyContextPointer);
                    string sbkey = referenceKeyManager.KeyToString(ref referenceKey);
                    surfaceBodyCounter++;
                    topologyMap.BodyIdentifiers["Body_" + surfaceBodyCounter] = sbkey;

                    foreach (Face face in body.Faces)
                    {
                        face.GetReferenceKey(ref referenceKey, keyContextPointer);
                        string key = referenceKeyManager.KeyToString(ref referenceKey);
                        faceCounter++;
                        topologyMap.FaceIdentifiers["Face_" + faceCounter] = key;
                    }
                    foreach (Edge edge in body.Edges)
                    {
                        if (edge.CurveType != CurveTypeEnum.kUnknownCurve)
                        {
                            edge.GetReferenceKey(ref referenceKey, keyContextPointer);
                            string key = referenceKeyManager.KeyToString(ref referenceKey);
                            edgeCounter++;
                            topologyMap.EdgeIdentifiers["Edge_" + edgeCounter] = key;
                        }
                    }
                }

                //Save key context
                byte[] keyContextArray = new byte[] { };
                referenceKeyManager.SaveContextToArray(keyContextPointer, ref keyContextArray);
                string keyContext = referenceKeyManager.KeyToString(ref keyContextArray);
                topologyMap.PartIdentifier = keyContext;
                topologyMap.PartNumber = ThePartDoc.DisplayName.Split('.')[0];
                JsonSerializerSettings settings = new JsonSerializerSettings
                {
                    PreserveReferencesHandling = PreserveReferencesHandling.Objects,
                    NullValueHandling = NullValueHandling.Ignore
                };
                var json = JsonConvert.SerializeObject(topologyMap, (Formatting)Formatting.Indented, settings);
                File.AppendAllText(Path + topologyMap.PartNumber + ".json", json);
            }

            MessageBox.Show("Data Collected", "Complete", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public bool IfFileExists()
        {
            if (ActiveDocument != null) return File.Exists(Path + ActiveDocument.DisplayName.Split('.')[0] + ".json");
            return false;
        }
    }
}
