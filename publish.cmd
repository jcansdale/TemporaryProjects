@set path="%VSInstallDir%\VSSDK\VisualStudioIntegration\Tools\Bin";%path%
VsixPublisher login -publisherName JamieCansdale -personalAccessToken %personalAccessToken%
VsixPublisher publish -payload TemporaryProjects.Vsix\bin\Debug\TemporaryProjects.vsix -publishManifest publishManifest.json
