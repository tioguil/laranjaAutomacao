import "dotenv/config";
import { chromium } from "playwright";
import { buildPunchLinesFromXlsx } from "./readHours";

function requireEnv(name: string): string {
    const value = process.env[name];
    if (!value || value.trim() === "") {
        throw new Error(`Missing required env var: ${name}`);
    }
    return value.trim();
}

async function main() {
    const codigoEmpregador = requireEnv("CODIGO_EMPREGADOR");
    const pin = requireEnv("PIN");
    const defaultEntry = requireEnv("DEFAULT_ENTRY");
    const filePath = requireEnv("FILE_PATH");
    const startDate = process.env.START_DATE?.trim() || undefined;

    const browser = await chromium.launch({ headless: false }); // true = sem janela
    const page = await browser.newPage();

    await page.goto("https://app.tangerino.com.br/Tangerino", { waitUntil: "domcontentloaded" });

    await page.locator('.login-abas li:nth-child(2) a').click();

    // Espere o input ficar visível
    await page.locator('input[name="codigoEmpregador"]').waitFor({ state: 'visible' });

    await page.locator('input[name="codigoEmpregador"]').fill(codigoEmpregador);
    await page.locator('input[name="pin"]').fill(pin);

    // Clique e espere a navegação acontecer (ou a URL mudar)
    await Promise.all([
        page.waitForNavigation({ waitUntil: "domcontentloaded" }),
        page.locator('input[name="btnLogin"]').click(),
    ]);

    await page.locator('header, aside, .menu-lateral').first().waitFor({
        timeout: 30_000
    });

    await page.waitForTimeout(2000);

    await page.locator('header a[rel="Registrar Ponto em Atraso"]').click();

    // Espere o input ficar visível
    await page.locator('input[name="dataPonto"]').waitFor({ state: 'visible' });

    await page.waitForTimeout(2000);

    // preenchendo o formulário
    const punches = buildPunchLinesFromXlsx(filePath, defaultEntry, startDate);
    console.table(punches);

    await page.pause();

    for (const punch of punches) {
        await page.locator('select[name="justificativa"]').selectOption('1'); // Esquecimento

        const dataInput = page.locator('input[name="dataPonto"]');

        await dataInput.click();              // ativa o onfocus (máscara)
        await dataInput.fill(punch.date);

        const horaInput = page.locator('input[name="horaPonto"]');

        await horaInput.click();              // ativa a máscara
        await horaInput.fill(punch.time);

        console.log(`${punch.date} ${punch.time} ${punch.type}`)

        await page.waitForTimeout(1000);

        await page.locator('input[name="saveAndContinue"]').click();
        await page.waitForTimeout(2000);
    }

    await page.pause();
    await browser.close();
}

main().catch((err) => {
    console.error("Erro na automação:", err);
    process.exit(1);
});


