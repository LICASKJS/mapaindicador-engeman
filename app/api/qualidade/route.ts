import { type NextRequest, NextResponse } from "next/server"
import * as XLSX from "xlsx"
import * as fs from "fs"
import * as path from "path"

export async function GET(request: NextRequest) {
  try {
    const filePath = path.join(process.cwd(), "dados", "atendimento controle_qualidade.xlsx")

    if (!fs.existsSync(filePath)) {
      return NextResponse.json({ error: "Arquivo de qualidade não encontrado" }, { status: 404 })
    }

    const fileBuffer = fs.readFileSync(filePath)
    const workbook = XLSX.read(fileBuffer, { type: "buffer" })
    const worksheet = workbook.Sheets[workbook.SheetNames[0]]
    const data = XLSX.utils.sheet_to_json(worksheet)

    return NextResponse.json({ qualidade: data })
  } catch (error) {
    console.error("[v0] Erro ao ler qualidade:", error)
    return NextResponse.json({ error: "Erro ao processar arquivo de qualidade" }, { status: 500 })
  }
}
