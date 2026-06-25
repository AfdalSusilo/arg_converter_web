import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'ARG Converter - CSV to Excel & Word',
  description: 'Convert CSV logs into formatted Excel and Word reports',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}
