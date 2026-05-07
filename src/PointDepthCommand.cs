using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using CivSurface = Autodesk.Civil.DatabaseServices.Surface;

namespace PointDepth;

public sealed class PointDepthExtension : IExtensionApplication
{
    public void Initialize()
    {
        Document? document = AcadApplication.DocumentManager.MdiActiveDocument;
        document?.Editor.WriteMessage("\nPointDepth loaded. Run AddPointDepth to write Depth_To_Surface values.");
    }

    public void Terminate()
    {
    }
}

public sealed class PointDepthCommand
{
    private const string CommandName = "AddPointDepth";
    private const string LegacyClassificationName = "PointDepth";
    private const string UnclassifiedClassificationName = "Unclassified";
    private const string UdpName = "Depth_To_Surface";
    private const int MaxSkippedDetails = 12;

    [CommandMethod(CommandName, CommandFlags.Modal)]
    public void AddPointDepth()
    {
        Document document = AcadApplication.DocumentManager.MdiActiveDocument;
        Editor editor = document.Editor;
        Database database = document.Database;
        CivilDocument civilDocument = CivilApplication.ActiveDocument;

        IReadOnlyList<PointGroupChoice> pointGroups = GetPointGroups(civilDocument, database);
        if (pointGroups.Count == 0)
        {
            editor.WriteMessage("\nPointDepth found no point groups in this drawing.");
            return;
        }

        PointGroupChoice? selectedPointGroup = PromptForPointGroup(editor, pointGroups);
        if (selectedPointGroup is null)
        {
            editor.WriteMessage("\nPointDepth canceled.");
            return;
        }

        ObjectId surfaceId = PromptForSurface(editor);
        if (surfaceId == ObjectId.Null)
        {
            editor.WriteMessage("\nPointDepth canceled.");
            return;
        }

        try
        {
            using (document.LockDocument())
            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
                PointGroup pointGroup = (PointGroup)transaction.GetObject(selectedPointGroup.ObjectId, OpenMode.ForWrite);
                CivSurface? surface = transaction.GetObject(surfaceId, OpenMode.ForRead) as CivSurface;
                if (surface is null)
                {
                    editor.WriteMessage("\nThe selected object is not a Civil 3D surface.");
                    return;
                }

                ReleaseLegacyClassification(pointGroup, civilDocument);
                UDPDouble depthUdp = GetOrCreateDepthUdp(civilDocument);

                uint[] pointNumbers = pointGroup.GetPointNumbers();
                if (pointNumbers.Length == 0)
                {
                    editor.WriteMessage($"\nPoint group \"{pointGroup.Name}\" contains no COGO points.");
                    return;
                }

                int updated = 0;
                int outsideSurface = 0;
                int failed = 0;
                List<string> skippedDetails = new();

                foreach (uint pointNumber in pointNumbers)
                {
                    try
                    {
                        ObjectId pointId = civilDocument.CogoPoints.GetPointByPointNumber(pointNumber);
                        CogoPoint point = (CogoPoint)transaction.GetObject(pointId, OpenMode.ForWrite);
                        double surfaceElevation = surface.FindElevationAtXY(point.Easting, point.Northing);
                        double depthToSurface = point.Elevation - surfaceElevation;

                        point.SetUDPValue(depthUdp, depthToSurface);
                        updated++;
                    }
                    catch (Autodesk.Civil.PointNotOnEntityException)
                    {
                        outsideSurface++;
                        AddSkippedDetail(skippedDetails, pointNumber, "outside selected surface");
                    }
                    catch (System.Exception ex)
                    {
                        failed++;
                        AddSkippedDetail(skippedDetails, pointNumber, ex.Message);
                    }
                }

                transaction.Commit();

                editor.WriteMessage(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "\nPointDepth wrote {0} {1} value(s) for point group \"{2}\" against surface \"{3}\".",
                        updated,
                        UdpName,
                        pointGroup.Name,
                        surface.Name));

                int skipped = outsideSurface + failed;
                if (skipped > 0)
                {
                    editor.WriteMessage(
                        string.Format(
                            CultureInfo.InvariantCulture,
                            "\nSkipped {0} point(s): {1} outside the surface, {2} failed for another reason.",
                            skipped,
                            outsideSurface,
                            failed));

                    foreach (string detail in skippedDetails)
                    {
                        editor.WriteMessage($"\n  {detail}");
                    }

                    if (skipped > skippedDetails.Count)
                    {
                        editor.WriteMessage($"\n  ... {skipped - skippedDetails.Count} more skipped point(s).");
                    }
                }
            }
        }
        catch (InvalidOperationException ex)
        {
            editor.WriteMessage($"\nPointDepth stopped: {ex.Message}");
        }
    }

    private static IReadOnlyList<PointGroupChoice> GetPointGroups(CivilDocument civilDocument, Database database)
    {
        List<PointGroupChoice> pointGroups = new();

        using Transaction transaction = database.TransactionManager.StartTransaction();
        int number = 1;
        foreach (ObjectId pointGroupId in civilDocument.PointGroups)
        {
            PointGroup pointGroup = (PointGroup)transaction.GetObject(pointGroupId, OpenMode.ForRead);
            pointGroups.Add(new PointGroupChoice(number, pointGroup.Name, pointGroup.PointsCount, pointGroupId));
            number++;
        }

        transaction.Commit();
        return pointGroups
            .OrderByDescending(group => string.Equals(group.Name, PointGroup.AllPointsGroupName, StringComparison.OrdinalIgnoreCase))
            .ThenBy(group => group.Name, StringComparer.OrdinalIgnoreCase)
            .Select((group, index) => group with { Number = index + 1 })
            .ToList();
    }

    private static PointGroupChoice? PromptForPointGroup(Editor editor, IReadOnlyList<PointGroupChoice> pointGroups)
    {
        editor.WriteMessage("\nPointDepth point groups:");
        foreach (PointGroupChoice group in pointGroups)
        {
            editor.WriteMessage(
                string.Format(
                    CultureInfo.InvariantCulture,
                    "\n  {0}. {1} ({2} point(s))",
                    group.Number,
                    group.Name,
                    group.PointCount));
        }

        PromptIntegerOptions options = new("\nSelect point group number: ")
        {
            AllowNegative = false,
            AllowNone = false,
            AllowZero = false,
            LowerLimit = 1,
            UpperLimit = pointGroups.Count
        };

        PromptIntegerResult result = editor.GetInteger(options);
        if (result.Status != PromptStatus.OK)
        {
            return null;
        }

        return pointGroups.First(group => group.Number == result.Value);
    }

    private static ObjectId PromptForSurface(Editor editor)
    {
        PromptEntityOptions options = new("\nSelect surface to compare point elevations against: ");
        options.SetRejectMessage("\nSelected object is not a Civil 3D surface.");
        options.AddAllowedClass(typeof(CivSurface), false);

        PromptEntityResult result = editor.GetEntity(options);
        return result.Status == PromptStatus.OK ? result.ObjectId : ObjectId.Null;
    }

    private static UDPDouble GetOrCreateDepthUdp(CivilDocument civilDocument)
    {
        UDP[] existingUdps = civilDocument.PointUDPs.ToArray();
        UDPDouble? existingDepthUdp = existingUdps
            .Where(udp => udp is not null && string.Equals(udp.Name, UdpName, StringComparison.OrdinalIgnoreCase))
            .OfType<UDPDouble>()
            .FirstOrDefault();

        if (existingDepthUdp is not null)
        {
            return existingDepthUdp;
        }

        UDP? conflictingUdp = existingUdps.FirstOrDefault(
            udp => udp is not null && string.Equals(udp.Name, UdpName, StringComparison.OrdinalIgnoreCase));
        if (conflictingUdp is not null)
        {
            throw new InvalidOperationException(
                $"A point UDP named {UdpName} already exists, but it is not a numeric double UDP.");
        }

        AttributeTypeInfoDouble typeInfo = new(UdpName)
        {
            Description = "PointDepth vertical difference: point elevation minus selected surface elevation. Positive is above the surface; negative is below.",
            UseDefaultValue = true,
            DefaultValue = 0.0,
            LowerBoundInclusive = true,
            LowerBoundValue = -1.0e9,
            UpperBoundInclusive = true,
            UpperBoundValue = 1.0e9
        };

        UDPClassification unclassified = GetOrCreateUnclassifiedClassification(civilDocument);
        return unclassified.CreateUDP(typeInfo);
    }

    private static UDPClassification GetOrCreateUnclassifiedClassification(CivilDocument civilDocument)
    {
        if (civilDocument.PointUDPClassifications.Contains(UnclassifiedClassificationName))
        {
            return civilDocument.PointUDPClassifications[UnclassifiedClassificationName];
        }

        foreach (UDPClassification classification in civilDocument.PointUDPClassifications)
        {
            if (string.IsNullOrWhiteSpace(classification.Name) ||
                string.Equals(classification.Name, UnclassifiedClassificationName, StringComparison.OrdinalIgnoreCase))
            {
                return classification;
            }
        }

        return civilDocument.PointUDPClassifications.Add(UnclassifiedClassificationName);
    }

    private static void ReleaseLegacyClassification(PointGroup pointGroup, CivilDocument civilDocument)
    {
        if (pointGroup.UDPClassificationApplyType == UDPClassificationApplyType.Custom &&
            string.Equals(pointGroup.UDPClassificationName, LegacyClassificationName, StringComparison.OrdinalIgnoreCase))
        {
            pointGroup.UseNoneClassification();
        }

        if (!civilDocument.PointUDPClassifications.Contains(LegacyClassificationName))
        {
            return;
        }

        UDPClassification legacyClassification = civilDocument.PointUDPClassifications[LegacyClassificationName];
        if (legacyClassification.UDPs.Count > 0)
        {
            return;
        }

        try
        {
            civilDocument.PointUDPClassifications.Remove(legacyClassification);
        }
        catch (InvalidOperationException)
        {
            // The old empty classification may still be assigned to another point group.
        }
    }

    private static void AddSkippedDetail(List<string> skippedDetails, uint pointNumber, string reason)
    {
        if (skippedDetails.Count >= MaxSkippedDetails)
        {
            return;
        }

        skippedDetails.Add(
            string.Format(
                CultureInfo.InvariantCulture,
                "Point {0}: {1}",
                pointNumber,
                reason));
    }

    private sealed record PointGroupChoice(int Number, string Name, uint PointCount, ObjectId ObjectId);
}
