"use client";

import { IExcelJSStyle, excelToJson, jsonToExcel } from "@/helpers/fileHandler";
import Image from "next/image";
import { ChangeEvent } from "react";

const data = [
  {
    bts_no: "S23C280002",
    store_code: "0281",
    supplier_code: "3WI01",
    supplier_name: "3 W INDONESIA PT",
    bts_date: "01 Jan 2023",
    termin_period: "01 Jun 2023 - 30 Jun 2023",
    contract_id: "123456789",
    status_code: "06",
    status_desc: "Waiting Finance Verification",
    trx_type: "AC AIR CURTAIN",
    taxinv: "test",
    taxinv_date: "16 Feb 2024",
    dpp: 9643000,
    ppn: 1060730,
    pph: 289290,
    ttf_no: "SBG0281-20240200002",
    wht_code: "K2",
    work_type: "Jasa Konstruksi",
    wht_rate: 3,
    taxinv_date_askdmakmdaksmdkmamdkasmdkamksdmkasmkasdmk_akmsdamdksmaksd: 10,
  },
];

export default function Home() {
  const onExportExcel = () => {
    const headerCells: IExcelJSStyle = {
      font: {
        bold: true,
      },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0000FF" },
      },
      alignment: {
        horizontal: "center",
      },
    };

    const dataCells: IExcelJSStyle = {
      font: {
        color: {
          argb: "FFFF0000",
        },
      },
    };

    const cells: IExcelJSStyle = {
      font: {
        size: 12,
      },
      border: {
        top: {
          style: "medium",
        },
        right: {
          style: "medium",
        },
        bottom: {
          style: "medium",
        },
        left: {
          style: "medium",
        },
      },
    };

    jsonToExcel({
      filename: "export_test_excel",
      data: {
        ["Sheet 1"]: {
          json: data,
          headerCells,
          dataCells,
          cells,
        },
        ["Sheet 2"]: {
          json: data,
        },
      },
    });
  };

  const onFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files && event.target.files[0];
    const jsonData = await excelToJson(file);
    console.log("jsonData", jsonData);
  };

  return (
    <main className="flex min-h-screen flex-col items-center justify-between p-24">
      <div className="z-10 max-w-5xl w-full items-center justify-between font-mono text-sm lg:flex">
        <p className="fixed left-0 top-0 flex w-full justify-center border-b border-gray-300 bg-gradient-to-b from-zinc-200 pb-6 pt-8 backdrop-blur-2xl dark:border-neutral-800 dark:bg-zinc-800/30 dark:from-inherit lg:static lg:w-auto  lg:rounded-xl lg:border lg:bg-gray-200 lg:p-4 lg:dark:bg-zinc-800/30">
          Get started by editing&nbsp;
          <code className="font-mono font-bold">app/page.tsx</code>
        </p>
        <div className="fixed bottom-0 left-0 flex h-48 w-full items-end justify-center bg-gradient-to-t from-white via-white dark:from-black dark:via-black lg:static lg:h-auto lg:w-auto lg:bg-none">
          <a
            className="pointer-events-none flex place-items-center gap-2 p-8 lg:pointer-events-auto lg:p-0"
            href="https://vercel.com?utm_source=create-next-app&utm_medium=appdir-template&utm_campaign=create-next-app"
            target="_blank"
            rel="noopener noreferrer"
          >
            By{" "}
            <Image
              src="/vercel.svg"
              alt="Vercel Logo"
              className="dark:invert"
              width={100}
              height={24}
              priority
            />
          </a>
        </div>
      </div>

      <div className="relative flex place-items-center before:absolute before:h-[300px] before:w-full sm:before:w-[480px] before:-translate-x-1/2 before:rounded-full before:bg-gradient-radial before:from-white before:to-transparent before:blur-2xl before:content-[''] after:absolute after:-z-20 after:h-[180px] after:w-full sm:after:w-[240px] after:translate-x-1/3 after:bg-gradient-conic after:from-sky-200 after:via-blue-200 after:blur-2xl after:content-[''] before:dark:bg-gradient-to-br before:dark:from-transparent before:dark:to-blue-700 before:dark:opacity-10 after:dark:from-sky-900 after:dark:via-[#0141ff] after:dark:opacity-40 before:lg:h-[360px] z-[-1]">
        <Image
          className="relative dark:drop-shadow-[0_0_0.3rem_#ffffff70] dark:invert"
          src="/next.svg"
          alt="Next.js Logo"
          width={180}
          height={37}
          priority
        />
      </div>

      <div className="mb-32 grid text-center lg:max-w-5xl lg:w-full lg:mb-0 lg:grid-cols-4 lg:text-left">
        <a
          href="https://nextjs.org/docs?utm_source=create-next-app&utm_medium=appdir-template&utm_campaign=create-next-app"
          className="group rounded-lg border border-transparent px-5 py-4 transition-colors hover:border-gray-300 hover:bg-gray-100 hover:dark:border-neutral-700 hover:dark:bg-neutral-800/30"
          target="_blank"
          rel="noopener noreferrer"
        >
          <h2 className={`mb-3 text-2xl font-semibold`}>
            Docs{" "}
            <span className="inline-block transition-transform group-hover:translate-x-1 motion-reduce:transform-none">
              -&gt;
            </span>
          </h2>
          <p className={`m-0 max-w-[30ch] text-sm opacity-50`}>
            Find in-depth information about Next.js features and API.
          </p>
        </a>

        <a
          href="https://nextjs.org/learn?utm_source=create-next-app&utm_medium=appdir-template-tw&utm_campaign=create-next-app"
          className="group rounded-lg border border-transparent px-5 py-4 transition-colors hover:border-gray-300 hover:bg-gray-100 hover:dark:border-neutral-700 hover:dark:bg-neutral-800/30"
          target="_blank"
          rel="noopener noreferrer"
        >
          <h2 className={`mb-3 text-2xl font-semibold`}>
            Learn{" "}
            <span className="inline-block transition-transform group-hover:translate-x-1 motion-reduce:transform-none">
              -&gt;
            </span>
          </h2>
          <p className={`m-0 max-w-[30ch] text-sm opacity-50`}>
            Learn about Next.js in an interactive course with&nbsp;quizzes!
          </p>
        </a>

        <a
          href="https://vercel.com/templates?framework=next.js&utm_source=create-next-app&utm_medium=appdir-template&utm_campaign=create-next-app"
          className="group rounded-lg border border-transparent px-5 py-4 transition-colors hover:border-gray-300 hover:bg-gray-100 hover:dark:border-neutral-700 hover:dark:bg-neutral-800/30"
          target="_blank"
          rel="noopener noreferrer"
        >
          <h2 className={`mb-3 text-2xl font-semibold`}>
            Templates{" "}
            <span className="inline-block transition-transform group-hover:translate-x-1 motion-reduce:transform-none">
              -&gt;
            </span>
          </h2>
          <p className={`m-0 max-w-[30ch] text-sm opacity-50`}>
            Explore starter templates for Next.js.
          </p>
        </a>

        <a
          href="https://vercel.com/new?utm_source=create-next-app&utm_medium=appdir-template&utm_campaign=create-next-app"
          className="group rounded-lg border border-transparent px-5 py-4 transition-colors hover:border-gray-300 hover:bg-gray-100 hover:dark:border-neutral-700 hover:dark:bg-neutral-800/30"
          target="_blank"
          rel="noopener noreferrer"
        >
          <h2 className={`mb-3 text-2xl font-semibold`}>
            Deploy{" "}
            <span className="inline-block transition-transform group-hover:translate-x-1 motion-reduce:transform-none">
              -&gt;
            </span>
          </h2>
          <p className={`m-0 max-w-[30ch] text-sm opacity-50 text-balance`}>
            Instantly deploy your Next.js site to a shareable URL with Vercel.
          </p>
        </a>
      </div>

      {/* Download Excel */}
      <button onClick={onExportExcel}>Export Excel</button>

      {/* Upload Excel */}
      <input type="file" name="fileName" id="fileId" onChange={onFileChange} />
    </main>
  );
}
