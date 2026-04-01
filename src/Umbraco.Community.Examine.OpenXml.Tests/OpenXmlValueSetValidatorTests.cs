using Examine;
using Umbraco.Cms.Core;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public class OpenXmlValueSetValidatorTests
{
    [Fact]
    public void Validate_WithValidPathAndNoParentId_ReturnsValid()
    {
        var validator = new OpenXmlValueSetValidator(false, null);
        var valueSet = CreateValueSet("1", "openxml", "File", "-1,1");

        var result = validator.Validate(valueSet);

        Assert.Equal(ValueSetValidationStatus.Valid, result.Status);
    }

    [Fact]
    public void Validate_MissingPath_ReturnsFailed()
    {
        var validator = new OpenXmlValueSetValidator(false, null);
        var values = new Dictionary<string, IEnumerable<object>>
        {
            ["nodeName"] = new object[] { "Test" }
        };
        var valueSet = new ValueSet("1", "openxml", "File", values);

        var result = validator.Validate(valueSet);

        Assert.Equal(ValueSetValidationStatus.Failed, result.Status);
    }

    [Fact]
    public void Validate_EmptyPath_ReturnsFailed()
    {
        var validator = new OpenXmlValueSetValidator(false, null);
        var valueSet = CreateValueSet("1", "openxml", "File", "");

        var result = validator.Validate(valueSet);

        Assert.Equal(ValueSetValidationStatus.Failed, result.Status);
    }

    [Fact]
    public void Validate_NullPathValue_ReturnsFailed()
    {
        var validator = new OpenXmlValueSetValidator(false, null);
        var values = new Dictionary<string, IEnumerable<object>>
        {
            ["path"] = new object[] { null! }
        };
        var valueSet = new ValueSet("1", "openxml", "File", values);

        var result = validator.Validate(valueSet);

        Assert.Equal(ValueSetValidationStatus.Failed, result.Status);
    }

    [Fact]
    public void ValidatePath_WithParentId_MatchingPath_ReturnsTrue()
    {
        var validator = new OpenXmlValueSetValidator(false, 100);

        var result = validator.ValidatePath("-1,100,200");

        Assert.True(result);
    }

    [Fact]
    public void ValidatePath_WithParentId_NonMatchingPath_ReturnsFalse()
    {
        var validator = new OpenXmlValueSetValidator(false, 100);

        var result = validator.ValidatePath("-1,200,300");

        Assert.False(result);
    }

    [Fact]
    public void ValidatePath_WithNoParentId_ReturnsTrue()
    {
        var validator = new OpenXmlValueSetValidator(false, null);

        var result = validator.ValidatePath("-1,1,2");

        Assert.True(result);
    }

    [Fact]
    public void ValidatePath_WithZeroParentId_ReturnsTrue()
    {
        var validator = new OpenXmlValueSetValidator(false, 0);

        var result = validator.ValidatePath("-1,1,2");

        Assert.True(result);
    }

    [Fact]
    public void ValidateRecycleBin_PublishedOnly_InMediaRecycleBin_ReturnsFalse()
    {
        var validator = new OpenXmlValueSetValidator(true, null);
        var recycleBinMediaId = Constants.System.RecycleBinMediaString;

        var result = validator.ValidateRecycleBin($"-1,{recycleBinMediaId},100", "media");

        Assert.False(result);
    }

    [Fact]
    public void ValidateRecycleBin_NotPublishedOnly_InRecycleBin_ReturnsTrue()
    {
        var validator = new OpenXmlValueSetValidator(false, null);
        var recycleBinMediaId = Constants.System.RecycleBinMediaString;

        var result = validator.ValidateRecycleBin($"-1,{recycleBinMediaId},100", "media");

        Assert.True(result);
    }

    [Fact]
    public void ValidateRecycleBin_PublishedOnly_NotInRecycleBin_ReturnsTrue()
    {
        var validator = new OpenXmlValueSetValidator(true, null);

        var result = validator.ValidateRecycleBin("-1,100,200", "media");

        Assert.True(result);
    }

    [Fact]
    public void Validate_InRecycleBin_PublishedOnly_ReturnsFiltered()
    {
        var validator = new OpenXmlValueSetValidator(true, null);
        var recycleBinMediaId = Constants.System.RecycleBinMediaString;
        var valueSet = CreateValueSet("1", "media", "File", $"-1,{recycleBinMediaId},100");

        var result = validator.Validate(valueSet);

        Assert.Equal(ValueSetValidationStatus.Filtered, result.Status);
    }

    [Fact]
    public void Validate_ParentIdNotInPath_ReturnsFiltered()
    {
        var validator = new OpenXmlValueSetValidator(false, 999);
        var valueSet = CreateValueSet("1", "openxml", "File", "-1,100,200");

        var result = validator.Validate(valueSet);

        Assert.Equal(ValueSetValidationStatus.Filtered, result.Status);
    }

    [Fact]
    public void Validate_ParentIdInPath_ReturnsValid()
    {
        var validator = new OpenXmlValueSetValidator(false, 100);
        var valueSet = CreateValueSet("1", "openxml", "File", "-1,100,200");

        var result = validator.Validate(valueSet);

        Assert.Equal(ValueSetValidationStatus.Valid, result.Status);
    }

    [Fact]
    public void Validate_WithPathInMediaRecycleBin_IsFiltered()
    {
        var validator = new OpenXmlValueSetValidator(true, null);
        var recycleBinMediaId = Constants.System.RecycleBinMediaString;
        // Specifically use the media recycle bin ID in the path
        var valueSet = CreateValueSet("1", "media", "File", $"-1,{recycleBinMediaId},200");

        var result = validator.Validate(valueSet);

        Assert.Equal(ValueSetValidationStatus.Filtered, result.Status);
    }

    [Fact]
    public void Validate_WithParentIdZero_DoesNotFilter()
    {
        // ParentId=0 should not filter because the code checks > 0
        var validator = new OpenXmlValueSetValidator(false, 0);
        var valueSet = CreateValueSet("1", "openxml", "File", "-1,100,200");

        var result = validator.Validate(valueSet);

        Assert.Equal(ValueSetValidationStatus.Valid, result.Status);
    }

    [Fact]
    public void Validate_ValueSetWithMultiplePathValues_UsesFirstPath()
    {
        // When multiple path values exist, the validator uses pathValues[0]
        var validator = new OpenXmlValueSetValidator(false, 100);
        var values = new Dictionary<string, IEnumerable<object>>
        {
            ["path"] = new object[] { "-1,100,200", "-1,999,888" }
        };
        var valueSet = new ValueSet("1", "openxml", "File", values);

        var result = validator.Validate(valueSet);

        // First path contains 100, so it should be valid
        Assert.Equal(ValueSetValidationStatus.Valid, result.Status);
    }

    private static ValueSet CreateValueSet(string id, string category, string itemType, string path)
    {
        var values = new Dictionary<string, IEnumerable<object>>
        {
            ["path"] = new object[] { path }
        };
        return new ValueSet(id, category, itemType, values);
    }
}
