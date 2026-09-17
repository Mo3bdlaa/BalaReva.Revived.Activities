# BalaReva Revived — Easy Image Activities

Image activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.EasyImage.Activities` 3.0.1, so
existing workflows keep binding after the swap. A test suite checks the surface back
against the original on every build.

This package targets `net8.0-windows` and that is not a style choice: the published
binding surface exposes `System.Drawing` types directly — `ImageRotate.FlipType` is a
`RotateFlipType`, and `ImageWatermark` takes a `Font`, a `Color`, a `Point` and an
`ImageFormat`. A workflow's `.xaml` names those types, so they have to stay, and
`System.Drawing.Common` is Windows-only from .NET 7 onward.

Activities: `ImageSizes`, `ImageResize`, `ImageCropper`, `ImageRotate`,
`ImageConverter`, `ImageCompression`, `ImageCombine`, `ImageMerge`, `ImageWatermark`.

See [docs/REVIVAL.md](https://github.com/Mo3bdlaa/BalaReva.Revived.Activities/blob/main/docs/REVIVAL.md).
