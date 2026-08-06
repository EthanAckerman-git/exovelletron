/**
 * Builds the Office Add-in manifest and installs it where Excel for Mac looks for
 * sideloaded add-ins: the `wef` folder inside Excel's sandbox container.
 *
 * Excel parses manifests once at launch, so changing the port means rewriting the
 * manifest and restarting Excel.
 */
import { writeFile, mkdir, rm, readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { paths as defaultPaths, ADDIN_ID, APP_NAME } from "../config.js";

export const MANIFEST_FILENAME = "excel-ai-local.manifest.xml";

const escapeXml = (s) =>
  String(s).replace(/[<>&"']/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", '"': "&quot;", "'": "&apos;" })[c]);

/**
 * @param {number} port  Port the local HTTPS server listens on.
 * @returns {string} manifest XML
 */
export function buildManifest(port) {
  if (!Number.isInteger(port) || port < 1024 || port > 65535) {
    throw new Error(`Invalid port for manifest: ${port}`);
  }
  const base = `https://localhost:${port}`;
  const desc =
    "Chat with a local AI model about your spreadsheet. Runs entirely on this Mac — " +
    "no cloud, no account, no data leaves your computer.";

  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>${ADDIN_ID}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>${escapeXml(APP_NAME)}</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="${escapeXml(APP_NAME)}"/>
  <Description DefaultValue="${escapeXml(desc)}"/>
  <IconUrl DefaultValue="${base}/assets/icon-64.png"/>
  <HighResolutionIconUrl DefaultValue="${base}/assets/icon-128.png"/>
  <SupportUrl DefaultValue="${base}/"/>
  <AppDomains>
    <AppDomain>${base}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="ExcelApi" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${base}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="ExcelAiLocal.Group">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="ExcelAiLocal.OpenPane">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ExcelAiLocalPane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${base}/assets/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="${base}/assets/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="${base}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${base}/taskpane.html"/>
        <bt:Url id="Commands.Url" DefaultValue="${base}/commands.html"/>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="${base}/"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="Local AI"/>
        <bt:String id="Button.Label" DefaultValue="${escapeXml(APP_NAME)}"/>
        <bt:String id="GetStarted.Title" DefaultValue="${escapeXml(APP_NAME)} is ready"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Button.Tooltip" DefaultValue="Open the offline AI assistant for this workbook"/>
        <bt:String id="GetStarted.Description" DefaultValue="Open the Home tab and choose ${escapeXml(APP_NAME)} to start."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}

export function manifestPath(p = defaultPaths) {
  return path.join(p.wefDir, MANIFEST_FILENAME);
}

/** True when an installed manifest exists and already points at `port`. */
export async function isManifestInstalled(port, p = defaultPaths) {
  const file = manifestPath(p);
  if (!existsSync(file)) return false;
  try {
    return (await readFile(file, "utf8")).includes(`https://localhost:${port}/taskpane.html`);
  } catch {
    return false;
  }
}

/** Write the manifest into Excel's sideload folder, creating it if absent. */
export async function installManifest(port, p = defaultPaths) {
  await mkdir(p.wefDir, { recursive: true });
  const file = manifestPath(p);
  await writeFile(file, buildManifest(port), "utf8");
  return file;
}

export async function uninstallManifest(p = defaultPaths) {
  await rm(manifestPath(p), { force: true });
}
