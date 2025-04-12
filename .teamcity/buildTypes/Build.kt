package _Self.buildTypes

import jetbrains.buildServer.configs.kotlin.*
import jetbrains.buildServer.configs.kotlin.buildFeatures.perfmon
import jetbrains.buildServer.configs.kotlin.buildSteps.DotnetMsBuildStep
import jetbrains.buildServer.configs.kotlin.buildSteps.dotnetBuild
import jetbrains.buildServer.configs.kotlin.buildSteps.dotnetMsBuild
import jetbrains.buildServer.configs.kotlin.buildSteps.nuGetInstaller
import jetbrains.buildServer.configs.kotlin.triggers.vcs

object Build : BuildType({
    name = "Build"

    vcs {
        root(DslContext.settingsRoot)
    }
steps {
    nuGetInstaller {
        id = "jb_nuget_installer"
        toolPath = "%teamcity.tool.NuGet.CommandLine.DEFAULT%"
        projects = "MoonDancer.sln"
        updatePackages = updateParams {
        }
    }
    dotnetBuild {
        id = "dotnet"
        projects = "MoonDancer.sln"
        sdk = "8"
    }
    dotnetMsBuild {
        id = "dotnet_1"
        projects = "MoonDancer.sln"
        version = DotnetMsBuildStep.MSBuildVersion.V17
        args = "-restore -noLogo"
        sdk = "8"
    }
}
    triggers {
        vcs {
        }
    }

    features {
        perfmon {
        }
    }
})