import * as React from 'react';

import {
  ITrendDatum
} from '../models/IDashboardChartData';

import {
  Panel
} from './BasePanel';

import styles from './ChartShared.module.scss';

export interface ISubmissionsTrendProps {
  items: readonly ITrendDatum[];

  onViewDetails?(): void;
}

interface ITrendPoint {
  label: string;
  value: number;
  x: number;
  y: number;
}

const SVG_WIDTH: number = 400;
const SVG_HEIGHT: number = 190;

const CHART_LEFT: number = 34;
const CHART_RIGHT: number = 382;
const CHART_TOP: number = 18;
const CHART_BOTTOM: number = 154;

const GRID_LINE_COUNT: number = 4;

export function SubmissionsTrend(
  props: ISubmissionsTrendProps
): React.ReactElement<ISubmissionsTrendProps> {
  const safeItems: ITrendDatum[] =
    props.items.map(
      (
        item: ITrendDatum
      ): ITrendDatum => {
        return {
          label:
            item.label,

          value:
            isFinite(item.value) &&
            !isNaN(item.value) &&
            item.value > 0
              ? item.value
              : 0
        };
      }
    );

  const maximumValue: number =
    getMaximumValue(
      safeItems
    );

  const yAxisMaximum: number =
    getRoundedAxisMaximum(
      maximumValue
    );

  const points: ITrendPoint[] =
    buildTrendPoints(
      safeItems,
      yAxisMaximum
    );

  const polylinePoints: string =
    points
      .map(
        (
          point: ITrendPoint
        ): string => {
          return (
            `${point.x.toString()},` +
            `${point.y.toString()}`
          );
        }
      )
      .join(' ');

  const gridLines: number[] =
    buildGridLines();

  return (
    <Panel
      title="Submissions Trend"
      subtitle="Number of submissions over time"
      onViewDetails={
        props.onViewDetails
      }
    >
      {points.length > 0 ? (
        <div className={styles.trendChart}>
          <svg
            className={styles.lineChart}
            viewBox={
              `0 0 ` +
              `${SVG_WIDTH.toString()} ` +
              `${SVG_HEIGHT.toString()}`
            }
            preserveAspectRatio="none"
            role="img"
            aria-label="Monthly supplier submission trend"
          >
            {gridLines.map(
              (
                yPosition: number,
                index: number
              ): React.ReactElement => {
                const percentageFromTop: number =
                  index /
                  GRID_LINE_COUNT;

                const gridValue: number =
                  Math.round(
                    yAxisMaximum *
                    (
                      1 -
                      percentageFromTop
                    )
                  );

                return (
                  <g
                    key={
                      `grid-${index.toString()}`
                    }
                  >
                    <line
                      x1={CHART_LEFT}
                      y1={yPosition}
                      x2={CHART_RIGHT}
                      y2={yPosition}
                      className={
                        styles.gridLine
                      }
                    />

                    <text
                      x={CHART_LEFT - 8}
                      y={yPosition + 3}
                      className={
                        styles.axisLabel
                      }
                      textAnchor="end"
                    >
                      {gridValue}
                    </text>
                  </g>
                );
              }
            )}

            <line
              x1={CHART_LEFT}
              y1={CHART_BOTTOM}
              x2={CHART_RIGHT}
              y2={CHART_BOTTOM}
              className={styles.axis}
            />

            {points.length > 1 && (
              <polyline
                points={polylinePoints}
                className={styles.line}
              />
            )}

            {points.map(
              (
                point: ITrendPoint
              ): React.ReactElement => {
                return (
                  <g key={point.label}>
                    <circle
                      cx={point.x}
                      cy={point.y}
                      r={7}
                      className={
                        styles.pointHitArea
                      }
                    >
                      <title>
                        {point.label}
                        {': '}
                        {point.value}
                        {' submissions'}
                      </title>
                    </circle>

                    <circle
                      cx={point.x}
                      cy={point.y}
                      r={4}
                      className={
                        styles.point
                      }
                    />

                    <text
                      x={point.x}
                      y={point.y - 10}
                      className={
                        styles.pointValue
                      }
                      textAnchor="middle"
                    >
                      {point.value}
                    </text>
                  </g>
                );
              }
            )}
          </svg>

          <div
            className={styles.labels}
            aria-hidden="true"
          >
            {points.map(
              (
                point: ITrendPoint
              ): React.ReactElement => {
                return (
                  <span
                    key={point.label}
                    title={point.label}
                  >
                    {point.label}
                  </span>
                );
              }
            )}
          </div>

          <div className={styles.trendLegend}>
            <span
              className={styles.trendLegendLine}
              aria-hidden="true"
            />

            <span>
              Submissions
            </span>
          </div>
        </div>
      ) : (
        <div
          className={styles.chartEmptyState}
          role="status"
        >
          <strong>
            No submission trend available
          </strong>

          <p>
            Monthly trend data will appear after
            questionnaire submissions are available.
          </p>
        </div>
      )}
    </Panel>
  );
}

function getMaximumValue(
  items: readonly ITrendDatum[]
): number {
  let maximumValue: number = 0;

  items.forEach(
    (
      item: ITrendDatum
    ): void => {
      if (
        item.value >
        maximumValue
      ) {
        maximumValue =
          item.value;
      }
    }
  );

  return Math.max(
    1,
    maximumValue
  );
}

function getRoundedAxisMaximum(
  maximumValue: number
): number {
  if (
    maximumValue <= 5
  ) {
    return 5;
  }

  if (
    maximumValue <= 10
  ) {
    return 10;
  }

  const interval: number =
    maximumValue <= 50
      ? 10
      : maximumValue <= 100
        ? 20
        : 50;

  return (
    Math.ceil(
      maximumValue /
      interval
    ) *
    interval
  );
}

function buildTrendPoints(
  items: readonly ITrendDatum[],
  maximumValue: number
): ITrendPoint[] {
  const chartWidth: number =
    CHART_RIGHT -
    CHART_LEFT;

  const chartHeight: number =
    CHART_BOTTOM -
    CHART_TOP;

  const denominator: number =
    Math.max(
      1,
      items.length - 1
    );

  return items.map(
    (
      item: ITrendDatum,
      index: number
    ): ITrendPoint => {
      const xPosition: number =
        items.length === 1
          ? (
            CHART_LEFT +
            chartWidth / 2
          )
          : (
            CHART_LEFT +
            (
              index /
              denominator
            ) *
            chartWidth
          );

      const valueRatio: number =
        maximumValue > 0
          ? item.value /
            maximumValue
          : 0;

      const yPosition: number =
        CHART_BOTTOM -
        valueRatio *
        chartHeight;

      return {
        label:
          item.label,

        value:
          item.value,

        x:
          Math.round(
            xPosition *
            100
          ) /
          100,

        y:
          Math.round(
            yPosition *
            100
          ) /
          100
      };
    }
  );
}

function buildGridLines(): number[] {
  const chartHeight: number =
    CHART_BOTTOM -
    CHART_TOP;

  const gridLines: number[] = [];

  for (
    let index: number = 0;
    index <= GRID_LINE_COUNT;
    index += 1
  ) {
    gridLines.push(
      CHART_TOP +
      (
        index /
        GRID_LINE_COUNT
      ) *
      chartHeight
    );
  }

  return gridLines;
}

export default SubmissionsTrend;