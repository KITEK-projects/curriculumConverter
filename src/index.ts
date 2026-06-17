import axios from 'axios'
import * as fs from 'fs'
import * as path from 'path'
import * as FormData from 'form-data'
import { app } from './parser'
import 'dotenv/config'
import { randomUUID } from 'crypto'

const VK_TOKEN = process.env.VK_TOKEN! // токен сообщества
const VK_GROUP_ID = process.env.VK_GROUP_ID! // id группы (без минуса)
const VK_API = 'https://api.vk.com/method'
const VK_VERSION = '5.199'

// ── утилиты VK API ──────────────────────────────────────────────────────────

async function vkCall<T = any>(
    method: string,
    params: Record<string, any>
): Promise<T> {
    const { data } = await axios.post(
        `${VK_API}/${method}`,
        new URLSearchParams({
            access_token: VK_TOKEN,
            v: VK_VERSION,
            ...Object.fromEntries(
                Object.entries(params).map(([k, v]) => [k, String(v)])
            ),
        }),
        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    )
    if (data.error)
        throw new Error(
            `VK API error ${data.error.error_code}: ${data.error.error_msg}`
        )
    return data.response
}

async function sendMessage(
    peerId: number,
    message: string,
    attachments?: string[]
) {
    await vkCall('messages.send', {
        peer_id: peerId,
        message,
        attachment: attachments?.join(',') ?? '',
        random_id: Date.now(),
    })
}

/** Скачать файл из VK и вернуть Buffer */
async function downloadVKFile(url: string): Promise<Buffer> {
    const response = await axios.get(url, { responseType: 'arraybuffer' })
    return Buffer.from(response.data)
}

/** Загрузить документ в VK и вернуть attachment-строку вида doc{owner}_{id} */
async function uploadDocument(
    filePath: string,
    filename: string,
    peerId: number
): Promise<string> {
    const { upload_url } = await vkCall<{ upload_url: string }>(
        'docs.getMessagesUploadServer',
        { type: 'doc', peer_id: peerId }
    )

    const form = new FormData()
    form.append('file', fs.createReadStream(filePath), { filename })
    const { data: uploadData } = await axios.post(upload_url, form, {
        headers: form.getHeaders(),
    })

    const saved = await vkCall<{
        type: string
        doc: { owner_id: number; id: number }
    }>('docs.save', { file: uploadData.file, title: filename })

    // VK возвращает { type: "doc", doc: { owner_id, id } }
    const doc = saved.doc
    return `doc${doc.owner_id}_${doc.id}`
}
// ── Long Poll ────────────────────────────────────────────────────────────────

interface LPServer {
    server: string
    key: string
    ts: string
}

async function getLPServer(): Promise<LPServer> {
    return vkCall<LPServer>('groups.getLongPollServer', {
        group_id: VK_GROUP_ID,
    })
}

interface LPResponse {
    ts: string
    updates: Array<{ type: string; object: any }>
    failed?: number
}

async function poll(
    server: string,
    key: string,
    ts: string
): Promise<LPResponse> {
    const { data } = await axios.get(server, {
        params: { act: 'a_check', key, ts, wait: 25 },
    })
    return data
}

// ── Обработка входящего сообщения ───────────────────────────────────────────

async function handleMessage(event: any) {
    const peerId: number = event.message.peer_id
    const fromId: number = event.message.from_id
    const attachments: any[] = event.message.attachments ?? []

    // Ищем вложение типа doc (документ / файл)
    const docAttachment = attachments.find((a: any) => a.type === 'doc')
    if (!docAttachment) return // не файл — игнорируем

    const doc = docAttachment.doc
    const url: string = doc.url
    const originalName: string = doc.title || 'input.xlsx'

    // Принимаем только xlsx
    if (!originalName.toLowerCase().endsWith('.xlsx')) {
        await sendMessage(peerId, 'Пожалуйста, пришли файл в формате .xlsx')
        return
    }

    console.log(`${new Date().toLocaleString()} | NEW QUERY | FROM: ${fromId}`)

    const sessionId = randomUUID()
    const baseDir = path.join(__dirname)
    const inputPath = path.join(baseDir, `input_${sessionId}.xlsx`)
    const outputPath = path.join(baseDir, `output_${sessionId}.xlsx`)
    const jsonPath = path.join(baseDir, `all.json`)

    try {
        // 1. Скачать файл
        const buffer = await downloadVKFile(url)
        fs.writeFileSync(inputPath, buffer)

        // 2. Обработать
        await app(inputPath, outputPath)

        // 3. Загрузить результат в VK и отправить
        const attachment = await uploadDocument(
            outputPath,
            `processed_${originalName}`,
            peerId
        )
        await sendMessage(peerId, 'Готово! Обработанный файл:', [attachment])
    } catch (error) {
        console.error(error)
        await sendMessage(peerId, 'Произошла ошибка при обработке файла.')
    } finally {
        if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath)
        if (fs.existsSync(jsonPath)) fs.unlinkSync(jsonPath)
        if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath)
    }
}

// ── Главный цикл ─────────────────────────────────────────────────────────────

async function main() {
    console.log('VK Bot is running...')
    let { server, key, ts } = await getLPServer()

    while (true) {
        try {
            const result = await poll(server, key, ts)

            if (result.failed) {
                // ts устарел или ключ сбросился — обновляем сервер
                const fresh = await getLPServer()
                server = fresh.server
                key = fresh.key
                ts = fresh.ts
                continue
            }

            ts = result.ts

            for (const update of result.updates) {
                if (update.type === 'message_new') {
                    handleMessage(update.object).catch(console.error)
                }
            }
        } catch (err) {
            console.error('Long poll error:', err)
            await new Promise((r) => setTimeout(r, 3000))
        }
    }
}

main()
