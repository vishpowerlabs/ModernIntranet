const avatarColors = [
    '#FFB900', '#FF8C00', '#F7630C', '#CA5010', '#DA3B01', '#EF6950', '#D13438', '#FF4343',
    '#E81123', '#EA4300', '#C239B3', '#E3008C', '#BF0077', '#C239B3', '#9A0089', '#0078D7',
    '#00B7C3', '#038387', '#00B294', '#018574', '#00CC6A', '#10893E', '#7A7574', '#5C2D91',
    '#008272', '#107C10', '#004B50', '#004B1C', '#32145A', '#2B579A', '#000000', '#102A4E'
];

export function stringToColor(str: string): string {
    if (!str) return avatarColors[0];
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = (str.codePointAt(i) || 0) + ((hash << 5) - hash);
        hash = Math.trunc(hash);
    }
    const index = Math.abs(hash) % avatarColors.length;
    return avatarColors[index];
}
