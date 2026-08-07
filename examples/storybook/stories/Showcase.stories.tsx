import type { Meta, StoryObj } from "@storybook/react-vite";
import { Spreadsheet } from "a1sheet/react";
import { salesReport } from "./fixtures.js";

/**
 * The demo on the landing page, and the one to look at first.
 *
 * Everything here is one workbook object: the currency and percent formats, the
 * frozen header, the total row, the conditional formatting on attainment, and
 * the `SORT` below that re-ranks the block. Nothing is configured on the
 * component — edit a booking figure and the total, the attainment percentages,
 * their colours, and the ranking all follow.
 */
const meta = {
  title: "Showcase/Q3 report",
  component: Spreadsheet,
  parameters: {
    layout: "fullscreen",
    docs: {
      description: {
        component:
          "A realistic sheet rather than a feature checklist. Same fixture the " +
          "landing page mounts, so what you see here is what ships.",
      },
    },
  },
} satisfies Meta<typeof Spreadsheet>;

export default meta;
type Story = StoryObj<typeof meta>;

export const Q3Report: Story = {
  name: "A quarter of regional numbers",
  args: { defaultWorkbook: salesReport(), height: 560 },
};
