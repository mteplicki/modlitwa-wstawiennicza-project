namespace Utils {
    const ONE_SECOND = 1000
    const ONE_MINUTE = ONE_SECOND * 60
    const MAX_EXECUTION_TIME = ONE_MINUTE * (5)
    export function isTimeLeft(START: number) : boolean {
      return MAX_EXECUTION_TIME > Date.now() - START;
    };
}

export default Utils;