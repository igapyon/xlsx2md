(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const markdownNormalizeHelper = requireXlsx2mdMarkdownNormalize();
    function renderNarrativeBlock(block) {
        if (!block.items || block.items.length === 0) {
            return block.lines.map((line) => markdownNormalizeHelper.normalizeMarkdownText(line)).join("\n");
        }
        const parts = [];
        let index = 0;
        while (index < block.items.length) {
            const current = block.items[index];
            const next = block.items[index + 1];
            if (current && next && next.startCol > current.startCol) {
                let childEnd = index + 1;
                while (childEnd < block.items.length && block.items[childEnd].startCol > current.startCol) {
                    childEnd += 1;
                }
                const childLines = block.items
                    .slice(index + 1, childEnd)
                    .map((item) => `- ${markdownNormalizeHelper.normalizeMarkdownListItemText(item.text)}`);
                parts.push(`### ${markdownNormalizeHelper.normalizeMarkdownHeadingText(current.text)}`);
                if (childLines.length > 0) {
                    parts.push(childLines.join("\n"));
                }
                index = childEnd;
                continue;
            }
            parts.push(markdownNormalizeHelper.normalizeMarkdownText(current.text));
            index += 1;
        }
        return parts.join("\n\n");
    }
    function isSectionHeadingNarrativeBlock(block) {
        if (!block || !block.items || block.items.length < 2) {
            return false;
        }
        return block.items[1].startCol > block.items[0].startCol;
    }
    const narrativeStructureApi = {
        renderNarrativeBlock,
        isSectionHeadingNarrativeBlock
    };
    moduleRegistry.registerModule("narrativeStructure", narrativeStructureApi);
})();
