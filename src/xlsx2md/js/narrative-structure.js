(() => {
    function renderNarrativeBlock(block) {
        if (!block.items || block.items.length === 0) {
            return block.lines.join("\n");
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
                    .map((item) => `- ${item.text}`);
                parts.push(`### ${current.text}`);
                if (childLines.length > 0) {
                    parts.push(childLines.join("\n"));
                }
                index = childEnd;
                continue;
            }
            parts.push(current.text);
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
    globalThis.__xlsx2mdNarrativeStructure = {
        renderNarrativeBlock,
        isSectionHeadingNarrativeBlock
    };
})();
