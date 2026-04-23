export const format = (template: string, ...args: any[]):string => {
    return template.replace(/{(\d+)}/g, (match, index) =>
        typeof args[index] !== 'undefined' ? args[index] : match
    );
}