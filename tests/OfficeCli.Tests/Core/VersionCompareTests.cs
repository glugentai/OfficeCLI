using FluentAssertions;
using Xunit;

namespace OfficeCli.Tests.Core;

public class VersionCompareTests
{
    private static bool IsNewer(string latest, string current)
    {
        var lp = latest.Split('.').Select(int.Parse).ToArray();
        var cp = current.Split('.').Select(int.Parse).ToArray();
        for (int i = 0; i < Math.Min(lp.Length, cp.Length); i++)
        {
            if (lp[i] > cp[i]) return true;
            if (lp[i] < cp[i]) return false;
        }
        return lp.Length > cp.Length;
    }

    [Theory]
    [InlineData("1.0.12", "1.0.11", true)]
    [InlineData("1.0.11", "1.0.12", false)]
    [InlineData("1.0.12", "1.0.12", false)]
    [InlineData("1.1.0", "1.0.11", true)]
    [InlineData("1.0.11", "1.1.0", false)]
    [InlineData("2.0.0", "1.9.9", true)]
    [InlineData("1.9.9", "2.0.0", false)]
    [InlineData("1.2.3", "1.2", true)]
    [InlineData("1.2", "1.2.3", false)]
    [InlineData("1.2", "1.2", false)]
    [InlineData("1.10.0", "1.9.0", true)]
    [InlineData("1.0.100", "1.0.99", true)]
    public void IsNewer_VersionComparison(string latest, string current, bool expected)
    {
        IsNewer(latest, current).Should().Be(expected);
    }
}
