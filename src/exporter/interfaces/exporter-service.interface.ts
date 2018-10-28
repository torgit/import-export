export interface IExporter<T, U> {
    export(input: T): Promise<U>
}