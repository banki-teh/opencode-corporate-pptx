import { tool } from "@opencode-ai/plugin"
import path from "path"
import fs from "fs"
import os from "os"

export default tool({
  description:
    "Generate a PowerPoint (PPTX) presentation in Banki.ru corporate style. " +
    "Accepts a JSON specification of slides and produces a .pptx file. " +
    "Supports slide types: title, section, content, two_column, metrics, image, thank_you. " +
    "Always load the 'corporate-pptx' skill first to understand the brand guidelines.",
  args: {
    slides_json: tool.schema
      .string()
      .describe(
        'JSON string with the presentation structure. Must contain "filename" (output path) and "slides" array. ' +
          "Each slide needs a \"type\" field (title|section|content|two_column|metrics|image|thank_you) " +
          "and type-specific fields like title, subtitle, bullets, metrics, etc. " +
          "See the corporate-pptx skill for the full JSON schema."
      ),
  },
  async execute(args, context) {
    const toolsDir = path.join(os.homedir(), ".config", "opencode", "tools")
    const scriptPath = path.join(toolsDir, "generate_pptx.py")

    if (!fs.existsSync(scriptPath)) {
      return `Error: Python generator script not found at ${scriptPath}`
    }

    // Parse and validate JSON
    let parsedData
    try {
      parsedData = JSON.parse(args.slides_json)
    } catch (e) {
      return `Error: Invalid JSON input: ${e}`
    }

    if (!parsedData.slides || !Array.isArray(parsedData.slides)) {
      return 'Error: JSON must contain a "slides" array'
    }

    // If filename is relative, make it relative to the working directory
    if (parsedData.filename && !path.isAbsolute(parsedData.filename)) {
      parsedData.filename = path.join(context.directory, parsedData.filename)
    } else if (!parsedData.filename) {
      parsedData.filename = path.join(context.directory, "presentation.pptx")
    }

    // Write JSON to temp file to avoid shell escaping issues
    const tmpFile = path.join(toolsDir, ".tmp_slides.json")
    fs.writeFileSync(tmpFile, JSON.stringify(parsedData), "utf-8")

    try {
      const result =
        await Bun.$`python3 ${scriptPath} --file ${tmpFile}`.text()
      return result.trim()
    } catch (e: any) {
      return `Error generating presentation: ${e.stderr || e.message || e}`
    } finally {
      // Clean up temp file
      try {
        fs.unlinkSync(tmpFile)
      } catch {}
    }
  },
})
