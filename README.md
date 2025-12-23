[![](https://dcbadge.limes.pink/api/server/https://discord.gg/hJJTrMNP5p?theme=default-inverted&style=for-the-badge)](https://discord.gg/hJJTrMNP5p)
# SysprepPreparator
The Sysprep preparation tool prepares computers for Sysprep generalization and for image capture.

> [!NOTE]
> This is currently in a separate repository so it can be worked on more easily. In the future, it will be merged into the DISMTools repository.

This tool is designed to be as modular as possible. Checks are performed by *Compatibility Checker Providers* (CCPs). Tasks are performed by *Preparation Tasks* (PTs).

*Special Thanks to [Real-MullaC](https://github.com/Real-MullaC) for helping with initial testing and expansion of this tool.*

## How do I get started?

No releases are available yet because this is still under construction. So you'll have to build it.

**Requirements:** Visual Studio 2012 and later, .NET Framework 4.8 Developer Pack

> [!NOTE]
> JetBrains Rider is supported, but you will not be able to design forms

1. Clone the repository
2. Open the solution and restore the NuGet packages (ManagedDism)
3. Press <kbd>Ctrl</kbd> + <kbd>F5</kbd> to launch the tool without a debugger

## Contributing to the tool

To do this:

1. Fork the repo
2. Do everything listed in the previous section, but clone your fork instead
3. Make your changes and test them
4. Create a pull request

> [!NOTE]
> Bonus points will be awarded to those who **DON'T VIBE CODE**

Additional documentation for registering Compatibility Checker Providers and Preparation Tasks will be mentioned now.

## Adding Compatibility Checker Providers and Preparation Tasks

Refer to [this documentation](./docs/CCP_PT/index.md) for more information.
