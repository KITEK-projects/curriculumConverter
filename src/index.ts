import { Telegraf } from "telegraf"
import axios from "axios"
import * as fs from "fs"
import * as path from "path"
import { app } from "./parser"
import "dotenv/config"
import { randomUUID } from "crypto"

const bot = new Telegraf(process.env.BOT_TOKEN)

bot.on("document", async (ctx) => {
    const fileId = ctx.message.document.file_id
    const fileName = ctx.message.document.file_name || "input.xlsx"
    const sessionId = randomUUID() // or Date.now().toString()

    const baseDir = path.join(__dirname) // your project root or a tmp folder
    const inputPath = path.join(baseDir, `input_${sessionId}.xlsx`)
    const outputPath = path.join(baseDir, `output_${sessionId}.xlsx`)
    // const inputPath = path.join(process.cwd(), `input_${sessionId}.xlsx`)
    // const outputPath = path.join(process.cwd(), `output_${sessionId}.xlsx`)

    try {
        // Download file
        const fileLink = await ctx.telegram.getFileLink(fileId)
        const response = await axios.get(fileLink.href, {
            responseType: "arraybuffer",
        })
        const inputBuffer = Buffer.from(response.data)
        fs.writeFileSync(inputPath, inputBuffer)

        // Pass paths to your parser
        await app(inputPath, outputPath)

        // Send result
        await ctx.replyWithDocument({
            source: outputPath,
            filename: `processed_${fileName}`,
        })
    } catch (error) {
        console.error(error)
        ctx.reply("There was an error processing your Excel file.")
    } finally {
        // Cleanup
        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath)
        if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath)
    }
})

bot.launch()
console.log("Bot is running...")
