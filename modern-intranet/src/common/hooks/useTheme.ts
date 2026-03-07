/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

import { ThemeService, IThemeTokens } from '../services/ThemeService';
import { useState, useEffect } from 'react';

export function useTheme(): IThemeTokens {
    const [theme, setTheme] = useState<IThemeTokens>(ThemeService.tokens);

    useEffect(() => {
        setTheme(ThemeService.tokens);
    }, []);

    return theme;
}
