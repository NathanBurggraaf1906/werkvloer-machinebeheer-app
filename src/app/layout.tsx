import type { Metadata } from "next";
import { Bellota_Text } from "next/font/google";
import "./globals.css";

const bellota = Bellota_Text({
  subsets: ["latin"],
  weight: ["300", "400", "700"],
  variable: "--font-bellota",
});

export const metadata: Metadata = {
  title: "Werkvloer Machinebeheer",
  description: "Mobiele werkvloer-app en beheerhub voor machinebeheer.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="nl" className={bellota.variable}>
      <head>
        <style
          id="brand-critical-theme"
          dangerouslySetInnerHTML={{
            __html: `
              body {
                background:
                  radial-gradient(circle at 12% 8%, rgba(225,6,0,.22), transparent 280px),
                  radial-gradient(circle at 88% 4%, rgba(162,144,97,.32), transparent 320px),
                  linear-gradient(180deg, #fff3f1, #f7f4ed 420px) !important;
              }
              .topbar {
                background: #fffdf8 !important;
                border: 2px solid #a29061 !important;
                border-bottom: 6px solid #a29061 !important;
              }
              .brandButton strong { color: #fff !important; }
              .brandMark {
                background: #e10600 !important;
                border: 2px solid #a29061 !important;
                color: #fff !important;
              }
              .mobileScreen {
                background: linear-gradient(180deg, #fffdf8 0%, #f6f0df 100%) !important;
                border: 2px solid #a29061 !important;
                border-top: 8px solid #000 !important;
              }
              .heroScreen {
                background:
                  linear-gradient(90deg, #000 0 18px, transparent 18px),
                  linear-gradient(180deg, rgba(162,144,97,.28), transparent 58%),
                  linear-gradient(135deg, transparent 0 70%, rgba(225,6,0,.16) 70% 100%),
                  #fffdf8 !important;
              }
              .heroCopy {
                background: #fff !important;
                border: 2px solid #a29061 !important;
                border-left: 8px solid #e10600 !important;
              }
              .actionButton.dark,
              .submitButton,
              .primaryButton {
                background: #000 !important;
                color: #fff !important;
                border-color: #a29061 !important;
              }
              .actionButton.light,
              .ghostButton,
              .smallButton,
              .iconButton {
                background: #fffdf8 !important;
                border: 2px solid #a29061 !important;
                color: #000 !important;
              }
              .topNavButton {
                background: #fff !important;
                border: 2px solid #a29061 !important;
                box-shadow: 0 5px 0 rgba(162,144,97,.22) !important;
                color: #000 !important;
              }
              .topNavButton.active {
                background: #000 !important;
                border-color: #a29061 !important;
                color: #fff !important;
              }
              .logoutButton {
                background: #e10600 !important;
                border: 2px solid #e10600 !important;
                color: #fff !important;
              }
              @media (max-width: 680px) {
                .topbar {
                  background: #fffdf8 !important;
                }
                .brandZone {
                  background: transparent !important;
                }
                .homeInBrand {
                  display: none !important;
                }
              }
            `,
          }}
        />
      </head>
      <body>{children}</body>
    </html>
  );
}
