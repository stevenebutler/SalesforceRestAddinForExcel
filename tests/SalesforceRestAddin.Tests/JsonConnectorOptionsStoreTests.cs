using System;
using System.IO;
using SalesforceRestAddin.Core.Session;

namespace SalesforceRestAddin.Tests;

public sealed class JsonConnectorOptionsStoreTests
{
    [Test]
    public async Task Load_MissingFile_ReturnsDefaults_Including_SkipHidden()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-options-{Guid.NewGuid():N}.json");
        var store = new JsonConnectorOptionsStore(path);

        var loaded = store.Load();

        await Assert.That(loaded.UseReference).IsFalse();
        await Assert.That(loaded.NoWarning).IsFalse();
        await Assert.That(loaded.NoConfirmQueryDownload).IsFalse();
        await Assert.That(loaded.NoQueryLimit).IsFalse();
        await Assert.That(loaded.AutoAssignRule).IsFalse();
        await Assert.That(loaded.IncludeHiddenCells).IsFalse();
        await Assert.That(loaded.ColumnSizingMode).IsEqualTo(ColumnSizingMode.FirstDownloadedPage);
        await Assert.That(loaded.RowSizingMode).IsEqualTo(RowSizingMode.ForceSingleLine);
        await Assert.That(loaded.CompositeBatchSize).IsEqualTo(200);
        await Assert.That(loaded.SendPreventAutoAssignHeader).IsTrue();
    }

    [Test]
    public async Task Load_EmptyObject_ReturnsDefaults_Including_SkipHidden()
    {
        var path = WriteTemp("{}");

        var loaded = new JsonConnectorOptionsStore(path).Load();

        await Assert.That(loaded.AutoAssignRule).IsFalse();
        await Assert.That(loaded.IncludeHiddenCells).IsFalse();
        await Assert.That(loaded.ColumnSizingMode).IsEqualTo(ColumnSizingMode.FirstDownloadedPage);
        await Assert.That(loaded.RowSizingMode).IsEqualTo(RowSizingMode.ForceSingleLine);

        File.Delete(path);
    }

    [Test]
    public async Task Load_OmittedValues_UseDefaults()
    {
        var path = WriteTemp(
            """
            {
              "useReference": true,
              "noWarning": false,
              "noQueryLimit": false,
              "autoAssignRule": false,
              "includeHiddenCells": false,
              "disableAutomaticSizing": true
            }
            """);

        var loaded = new JsonConnectorOptionsStore(path).Load();

        await Assert.That(loaded.UseReference).IsTrue();
        await Assert.That(loaded.NoWarning).IsFalse();
        await Assert.That(loaded.NoQueryLimit).IsFalse();
        await Assert.That(loaded.AutoAssignRule).IsFalse();
        await Assert.That(loaded.IncludeHiddenCells).IsFalse();
        await Assert.That(loaded.ColumnSizingMode).IsEqualTo(ColumnSizingMode.FirstDownloadedPage);
        await Assert.That(loaded.RowSizingMode).IsEqualTo(RowSizingMode.ForceSingleLine);

        File.Delete(path);
    }

    [Test]
    public async Task Load_LegacySkipHiddenCellsTrue_MigratesToIncludeHiddenFalse()
    {
        var path = WriteTemp(
            """
            {
              "skipHiddenCells": true
            }
            """);

        var loaded = new JsonConnectorOptionsStore(path).Load();

        await Assert.That(loaded.IncludeHiddenCells).IsFalse();

        File.Delete(path);
    }

    [Test]
    public async Task Load_LegacySkipHiddenCellsFalse_MigratesToIncludeHiddenTrue()
    {
        var path = WriteTemp(
            """
            {
              "skipHiddenCells": false
            }
            """);

        var loaded = new JsonConnectorOptionsStore(path).Load();

        await Assert.That(loaded.IncludeHiddenCells).IsTrue();

        File.Delete(path);
    }

    [Test]
    public async Task Load_IncludeHiddenCells_TakesPrecedenceOverLegacySkipKey()
    {
        var path = WriteTemp(
            """
            {
              "includeHiddenCells": false,
              "skipHiddenCells": false
            }
            """);

        var loaded = new JsonConnectorOptionsStore(path).Load();

        await Assert.That(loaded.IncludeHiddenCells).IsFalse();

        File.Delete(path);
    }

    [Test]
    public async Task Load_InvalidJson_ReturnsAllDefaults_And_LogsWarning()
    {
        var path = WriteTemp("not-json");
        var warnings = new List<string>();

        var loaded = new JsonConnectorOptionsStore(path, warnings.Add).Load();

        await Assert.That(loaded.UseReference).IsFalse();
        await Assert.That(loaded.ColumnSizingMode).IsEqualTo(ColumnSizingMode.FirstDownloadedPage);
        await Assert.That(loaded.RowSizingMode).IsEqualTo(RowSizingMode.ForceSingleLine);
        await Assert.That(warnings.Count).IsEqualTo(1);
        await Assert.That(warnings[0]).Contains("settings JSON is invalid");

        File.Delete(path);
    }

    [Test]
    public async Task SaveAndLoad_RoundTripsAllFlags()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-options-{Guid.NewGuid():N}.json");
        var store = new JsonConnectorOptionsStore(path);
        var saved = new ConnectorOptions
        {
            UseReference = true,
            NoWarning = true,
            NoConfirmQueryDownload = true,
            NoQueryLimit = true,
            AutoAssignRule = true,
            IncludeHiddenCells = true,
            ColumnSizingMode = ColumnSizingMode.HeadersOnly,
            RowSizingMode = RowSizingMode.ForceSingleLine,
        };

        store.Save(saved);
        var loaded = store.Load();

        await Assert.That(loaded.UseReference).IsTrue();
        await Assert.That(loaded.NoWarning).IsTrue();
        await Assert.That(loaded.NoConfirmQueryDownload).IsTrue();
        await Assert.That(loaded.NoQueryLimit).IsTrue();
        await Assert.That(loaded.AutoAssignRule).IsTrue();
        await Assert.That(loaded.IncludeHiddenCells).IsTrue();
        await Assert.That(loaded.ColumnSizingMode).IsEqualTo(ColumnSizingMode.HeadersOnly);
        await Assert.That(loaded.RowSizingMode).IsEqualTo(RowSizingMode.ForceSingleLine);
        await Assert.That(loaded.SendPreventAutoAssignHeader).IsFalse();

        File.Delete(path);
    }

    [Test]
    public async Task Save_PersistsExplicitFalse_ForAutoAssignRule()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-options-{Guid.NewGuid():N}.json");
        var store = new JsonConnectorOptionsStore(path);

        store.Save(new ConnectorOptions { AutoAssignRule = false });
        var loaded = store.Load();

        await Assert.That(loaded.AutoAssignRule).IsFalse();
        await Assert.That(loaded.SendPreventAutoAssignHeader).IsTrue();

        File.Delete(path);
    }

    [Test]
    public async Task Load_MissingSizingModes_UsesDefaults_And_IgnoresLegacySizingSetting()
    {
        var path = WriteTemp(
            """
            {
              "disableAutomaticSizing": true
            }
            """);

        var loaded = new JsonConnectorOptionsStore(path).Load();

        await Assert.That(loaded.ColumnSizingMode).IsEqualTo(ColumnSizingMode.FirstDownloadedPage);
        await Assert.That(loaded.RowSizingMode).IsEqualTo(RowSizingMode.ForceSingleLine);

        File.Delete(path);
    }

    [Test]
    public async Task Load_InvalidSizingModes_DefaultsOnlyThoseSettings_And_LogsOneWarning()
    {
        var path = WriteTemp(
            """
            {
              "useReference": true,
              "columnSizingMode": "futureMode",
              "rowSizingMode": 2
            }
            """);
        var warnings = new List<string>();

        var loaded = new JsonConnectorOptionsStore(path, warnings.Add).Load();

        await Assert.That(loaded.UseReference).IsTrue();
        await Assert.That(loaded.ColumnSizingMode).IsEqualTo(ColumnSizingMode.FirstDownloadedPage);
        await Assert.That(loaded.RowSizingMode).IsEqualTo(RowSizingMode.ForceSingleLine);
        await Assert.That(warnings.Count).IsEqualTo(1);
        await Assert.That(warnings[0]).Contains("columnSizingMode, rowSizingMode");

        File.Delete(path);
    }

    [Test]
    public async Task T_OPT_03_CompositeBatchSize_Persisted_And_Used_By_Splitter()
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-options-{Guid.NewGuid():N}.json");
        var store = new JsonConnectorOptionsStore(path);
        store.Save(new ConnectorOptions { CompositeBatchSize = 50 });
        var loaded = store.Load();

        await Assert.That(loaded.CompositeBatchSize).IsEqualTo(50);

        var batches = SalesforceRestAddin.Core.DataPlane.RecordBatchSplitter.Split(
            Enumerable.Range(0, 120).Select(i => i.ToString()).ToList(),
            loaded.CompositeBatchSize);
        await Assert.That(batches.Count).IsEqualTo(3);

        File.Delete(path);
    }

    private static string WriteTemp(string contents)
    {
        var path = Path.Combine(Path.GetTempPath(), $"forceconnector-options-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, contents);
        return path;
    }
}
