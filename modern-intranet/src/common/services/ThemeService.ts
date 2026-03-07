import { useEffect, useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IThemeTokens {
    themePrimary: string;
    themeDark: string;
    themeDarker: string;
    themeSecondary: string;
    themeLighter: string;
    themeLighterAlt: string;
    neutralPrimary: string;
    neutralSecondary: string;
    neutralTertiary: string;
    neutralLight: string;
    neutralLighter: string;
    neutralLighterAlt: string;
    white: string;
    black: string;
}

const defaultTheme: IThemeTokens = {
    themePrimary: '#0078d4',
    themeDark: '#005a9e',
    themeDarker: '#004578',
    themeSecondary: '#2b88d8',
    themeLighter: '#deecf9',
    themeLighterAlt: '#eff6fc',
    neutralPrimary: '#323130',
    neutralSecondary: '#605e5c',
    neutralTertiary: '#a19f9d',
    neutralLight: '#edebe9',
    neutralLighter: '#f3f2f1',
    neutralLighterAlt: '#faf9f8',
    white: '#ffffff',
    black: '#000000'
};

export class ThemeService {
    private static _tokens: IThemeTokens = { ...defaultTheme };

    public static initialize(context: WebPartContext): void {
        let themeInfo = null;

        // Check Teams context if available
        const teamsContext = (context as any)?.sdks?.microsoftTeams?.teamsJs?.context;
        if (teamsContext?.theme) {
            // In a real app we'd map Teams theme to our tokens here based on the theme string, 
            // but typically we rely on SharePoint theme state for consistent SPO look
        }

        // Read from window.__themeState__
        const win = globalThis as any;
        if (win?.__themeState__?.theme) {
            themeInfo = win.__themeState__.theme;
        }

        if (themeInfo) {
            this._tokens = {
                themePrimary: themeInfo.themePrimary || defaultTheme.themePrimary,
                themeDark: themeInfo.themeDark || defaultTheme.themeDark,
                themeDarker: themeInfo.themeDarker || defaultTheme.themeDarker,
                themeSecondary: themeInfo.themeSecondary || defaultTheme.themeSecondary,
                themeLighter: themeInfo.themeLighter || defaultTheme.themeLighter,
                themeLighterAlt: themeInfo.themeLighterAlt || defaultTheme.themeLighterAlt,
                neutralPrimary: themeInfo.neutralPrimary || defaultTheme.neutralPrimary,
                neutralSecondary: themeInfo.neutralSecondary || defaultTheme.neutralSecondary,
                neutralTertiary: themeInfo.neutralTertiary || defaultTheme.neutralTertiary,
                neutralLight: themeInfo.neutralLight || defaultTheme.neutralLight,
                neutralLighter: themeInfo.neutralLighter || defaultTheme.neutralLighter,
                neutralLighterAlt: themeInfo.neutralLighterAlt || defaultTheme.neutralLighterAlt,
                white: themeInfo.white || defaultTheme.white,
                black: themeInfo.black || defaultTheme.black
            };
        }
    }

    public static get tokens(): IThemeTokens {
        return this._tokens;
    }

    public static getThemeCSS(): string {
        return `
      --themePrimary: ${this._tokens.themePrimary};
      --themeDark: ${this._tokens.themeDark};
      --themeDarker: ${this._tokens.themeDarker};
      --themeSecondary: ${this._tokens.themeSecondary};
      --themeLighter: ${this._tokens.themeLighter};
      --themeLighterAlt: ${this._tokens.themeLighterAlt};
      --neutralPrimary: ${this._tokens.neutralPrimary};
      --neutralSecondary: ${this._tokens.neutralSecondary};
      --neutralTertiary: ${this._tokens.neutralTertiary};
      --neutralLight: ${this._tokens.neutralLight};
      --neutralLighter: ${this._tokens.neutralLighter};
      --neutralLighterAlt: ${this._tokens.neutralLighterAlt};
      --white: ${this._tokens.white};
      --black: ${this._tokens.black};
    `;
    }
}

export function useTheme(): IThemeTokens {
    const [theme, setTheme] = useState<IThemeTokens>(ThemeService.tokens);

    useEffect(() => {
        // In a fully dynamic scenario, we might listen to theme change events here.
        // For now, returning the tokens statically from the singleton Service.
        setTheme(ThemeService.tokens);
    }, []);

    return theme;
}
