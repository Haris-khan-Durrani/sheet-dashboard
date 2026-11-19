import { NextResponse } from 'next/server';
import { handleOAuthCallback } from '../../../../lib/sheets-engine';

export const dynamic = 'force-dynamic';

export async function GET(request: Request) {
  try {
    const { searchParams, origin } = new URL(request.url);
    const code = searchParams.get('code');
    if (!code) {
      // If user opened /auth/google/callback directly without a code, send them to /auth/google to start OAuth
      return NextResponse.redirect(new URL('/auth/google', origin));
    }
    await handleOAuthCallback(code);
    return NextResponse.redirect(new URL('/', origin));
  } catch (err: any) {
    return new NextResponse('Google OAuth Error: ' + err.message, { status: 500 });
  }
}
