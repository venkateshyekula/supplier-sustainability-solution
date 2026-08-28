import * as React from 'react';

import {
  IChartDatum
} from '../models/IDashboardChartData';

import {
  Panel
} from './BasePanel';

import styles from './ChartShared.module.scss';

export interface IDonutPanelProps {
  title: string;
  subtitle?: string;

  items: readonly IChartDatum[];

  onViewDetails?(): void;
}

interface IDonutSegment {
  key: string;
  label: string;
  value: number;
  percentage: number;
  color: string;

  startAngle: number;
  endAngle: number;
  middleAngle: number;
}

interface IPoint {
  x: number;
  y: number;
}

const VIEWBOX_SIZE: number = 200;
const CENTER: number = VIEWBOX_SIZE / 2;

const OUTER_RADIUS: number = 82;
const INNER_RADIUS: number = 44;
const LABEL_RADIUS: number = 64;

const START_ANGLE: number = -90;

export function DonutPanel(
  props: IDonutPanelProps
): React.ReactElement<IDonutPanelProps> {
  const total: number =
    props.items.reduce(
      (
        runningTotal: number,
        item: IChartDatum
      ): number => {
        const safeValue: number =
          isFinite(item.value) &&
          !isNaN(item.value) &&
          item.value > 0
            ? item.value
            : 0;

        return runningTotal + safeValue;
      },
      0
    );

  const segments: IDonutSegment[] =
    buildSegments(
      props.items,
      total
    );

  return (
    <Panel
      title={props.title}
      subtitle={props.subtitle}
      onViewDetails={props.onViewDetails}
    >
      {total > 0 ? (
        <div className={styles.donutArea}>
          <div className={styles.donutChart}>
            <svg
              className={styles.donutSvg}
              viewBox={
                `0 0 ` +
                `${VIEWBOX_SIZE.toString()} ` +
                `${VIEWBOX_SIZE.toString()}`
              }
              role="img"
              aria-label={
                `${props.title}. ` +
                `Total ${total.toString()}`
              }
            >
              <circle
                className={styles.donutBackground}
                cx={CENTER}
                cy={CENTER}
                r={
                  (
                    OUTER_RADIUS +
                    INNER_RADIUS
                  ) / 2
                }
                fill="none"
                strokeWidth={
                  OUTER_RADIUS -
                  INNER_RADIUS
                }
              />

              {segments.map(
                (
                  segment: IDonutSegment
                ): React.ReactElement => {
                  const labelPoint: IPoint =
                    polarToCartesian(
                      CENTER,
                      CENTER,
                      LABEL_RADIUS,
                      segment.middleAngle
                    );

                  /*
                   * Hide labels for extremely small segments
                   * because the text will not fit inside.
                   */
                  const showPercentage: boolean =
                    segment.percentage >= 5;

                  return (
                    <g key={segment.key}>
                      <path
                        className={styles.donutSegment}
                        d={
                          describeDonutSegment(
                            CENTER,
                            CENTER,
                            OUTER_RADIUS,
                            INNER_RADIUS,
                            segment.startAngle,
                            segment.endAngle
                          )
                        }
                        fill={segment.color}
                      >
                        <title>
                          {segment.label}
                          {': '}
                          {segment.value.toString()}
                          {' ('}
                          {segment.percentage.toString()}
                          {'%)'}
                        </title>
                      </path>

                      {showPercentage && (
                        <text
                          className={
                            styles.donutPercentage
                          }
                          x={labelPoint.x}
                          y={labelPoint.y}
                          textAnchor="middle"
                          dominantBaseline="middle"
                          aria-hidden="true"
                        >
                          {segment.percentage}%
                        </text>
                      )}
                    </g>
                  );
                }
              )}

              <circle
                className={styles.donutCenterCircle}
                cx={CENTER}
                cy={CENTER}
                r={INNER_RADIUS - 1}
              />
            </svg>
          </div>

          <div className={styles.legend}>
            {segments.map(
              (
                segment: IDonutSegment
              ): React.ReactElement => {
                return (
                  <div
                    key={segment.key}
                    className={styles.legendRow}
                  >
                    <span
                      className={styles.dot}
                      style={{
                        backgroundColor:
                          segment.color
                      }}
                      aria-hidden="true"
                    />

                    <span
                      className={
                        styles.legendLabel
                      }
                      title={segment.label}
                    >
                      {segment.label}
                    </span>

                    <strong
                      className={
                        styles.legendValue
                      }
                    >
                      {segment.value}
                    </strong>

                    <span
                      className={
                        styles.legendPercentage
                      }
                    >
                      {segment.percentage}%
                    </span>
                  </div>
                );
              }
            )}

            <div className={styles.legendTotal}>
              <span>Total</span>

              <strong>
                {total}
              </strong>
            </div>
          </div>
        </div>
      ) : (
        <div
          className={styles.chartEmptyState}
          role="status"
        >
          <strong>
            No chart data available
          </strong>

          <p>
            Chart values will appear after ESG
            submission data has been loaded.
          </p>
        </div>
      )}
    </Panel>
  );
}

function buildSegments(
  items: readonly IChartDatum[],
  total: number
): IDonutSegment[] {
  let currentAngle: number =
    START_ANGLE;

  return items
    .filter(
      (
        item: IChartDatum
      ): boolean => {
        return (
          isFinite(item.value) &&
          !isNaN(item.value) &&
          item.value > 0
        );
      }
    )
    .map(
      (
        item: IChartDatum,
        index: number
      ): IDonutSegment => {
        const percentage: number =
          total > 0
            ? Math.round(
                (
                  item.value /
                  total
                ) *
                100
              )
            : 0;

        const segmentAngle: number =
          total > 0
            ? (
                item.value /
                total
              ) *
              360
            : 0;

        const startAngle: number =
          currentAngle;

        const endAngle: number =
          currentAngle +
          segmentAngle;

        const middleAngle: number =
          startAngle +
          segmentAngle / 2;

        currentAngle =
          endAngle;

        return {
          key:
            `${item.label}-${index.toString()}`,

          label:
            item.label,

          value:
            item.value,

          percentage,

          color:
            item.color,

          startAngle,

          endAngle,

          middleAngle
        };
      }
    );
}

function describeDonutSegment(
  centerX: number,
  centerY: number,
  outerRadius: number,
  innerRadius: number,
  startAngle: number,
  endAngle: number
): string {
  /*
   * SVG cannot produce a complete circle from a single
   * arc command. Keep full-circle segments slightly below
   * 360 degrees.
   */
  const safeEndAngle: number =
    endAngle - startAngle >= 360
      ? startAngle + 359.999
      : endAngle;

  const outerStart: IPoint =
    polarToCartesian(
      centerX,
      centerY,
      outerRadius,
      startAngle
    );

  const outerEnd: IPoint =
    polarToCartesian(
      centerX,
      centerY,
      outerRadius,
      safeEndAngle
    );

  const innerEnd: IPoint =
    polarToCartesian(
      centerX,
      centerY,
      innerRadius,
      safeEndAngle
    );

  const innerStart: IPoint =
    polarToCartesian(
      centerX,
      centerY,
      innerRadius,
      startAngle
    );

  const largeArcFlag: number =
    safeEndAngle -
      startAngle >
    180
      ? 1
      : 0;

  return [
    `M ${outerStart.x} ${outerStart.y}`,

    `A ${outerRadius} ${outerRadius} ` +
      `0 ${largeArcFlag} 1 ` +
      `${outerEnd.x} ${outerEnd.y}`,

    `L ${innerEnd.x} ${innerEnd.y}`,

    `A ${innerRadius} ${innerRadius} ` +
      `0 ${largeArcFlag} 0 ` +
      `${innerStart.x} ${innerStart.y}`,

    'Z'
  ].join(' ');
}

function polarToCartesian(
  centerX: number,
  centerY: number,
  radius: number,
  angleInDegrees: number
): IPoint {
  const angleInRadians: number =
    angleInDegrees *
    Math.PI /
    180;

  return {
    x:
      Math.round(
        (
          centerX +
          radius *
          Math.cos(
            angleInRadians
          )
        ) *
        100
      ) /
      100,

    y:
      Math.round(
        (
          centerY +
          radius *
          Math.sin(
            angleInRadians
          )
        ) *
        100
      ) /
      100
  };
}

export default DonutPanel;